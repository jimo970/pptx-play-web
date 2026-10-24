<script setup lang="ts">
import type { PptxDocument, PptxHyperlink, PptxSlide, PptxTransitionPlayback, PptxTriggerPlayback } from '~/utils/pptx-model'
import type { LocaleMessageKey } from '~/composables/useLocale'
import { disposePptxDocument, parsePptxInWorker, stopPptxParserWorker } from '~/utils/pptx-parser-client'
import { presenterChannelName, type PresenterMessage } from '~/utils/presenter-channel'
import { useSlideTransition } from '~/composables/useSlideTransition'

const { locale, t, toggleLocale, translateError, translateWarning } = useLocale()

const fileInput = ref<HTMLInputElement | null>(null)
let fileLoadGeneration = 0
const deck = shallowRef<PptxDocument | null>(null)
const deckGeneration = ref(0)
const currentIndex = ref(0)
const animationStep = ref(0)
const errorMessage = ref('')
const isLoading = ref(false)
const isDragging = ref(false)
const presenterView = ref(false)
const sorterView = ref(false)
const readingView = ref(false)
const readingNotesOpen = ref(false)
const presenterElapsedSeconds = ref(0)
const projectorConnected = ref(false)
const projectorOpen = ref(false)
const projectorError = ref<LocaleMessageKey | ''>('')
const projectorAcknowledgement = ref<{ slideId: string; number: number; animationStep: number; blank: 'black' | 'white' | null; triggerCount: number } | null>(null)
const triggerPlayback = shallowRef<PptxTriggerPlayback[]>([])
const slideTimelineReadyKey = ref('')
const transitionPlayback = shallowRef<PptxTransitionPlayback | null>(null)
const transitionAudio = ref<HTMLAudioElement | null>(null)
const transitionSoundWarning = ref('')
let transitionSoundPlaybackId = 0
let transitionSoundShouldPlay = false
const presenterTrigger = ref<HTMLButtonElement | null>(null)
const presenterExit = ref<HTMLButtonElement | null>(null)
const sorterTrigger = ref<HTMLButtonElement | null>(null)
const sorterExit = ref<HTMLButtonElement | null>(null)
const readingTrigger = ref<HTMLButtonElement | null>(null)
const readingExit = ref<HTMLButtonElement | null>(null)
const draggedSlideIndex = ref<number | null>(null)
const sorterPageStart = ref(0)
const screenOverlay = ref<'black' | 'white' | null>(null)
const thumbnailList = ref<HTMLElement | null>(null)
const thumbnailStart = ref(0)
const thumbnailEnd = ref(12)
const thumbnailStep = ref(156)
const thumbnailGap = ref(10)
const thumbnailAxis = ref<'vertical' | 'horizontal'>('vertical')
let thumbnailFrame = 0
const stageShell = ref<HTMLElement | null>(null)
const stageSlideHost = ref<HTMLElement | null>(null)
const transitionUnderlay = ref<HTMLElement | null>(null)
const flashOverlay = ref<HTMLElement | null>(null)
const presenterStageShell = ref<HTMLElement | null>(null)
const presenterSlideHost = ref<HTMLElement | null>(null)
const presenterTransitionUnderlay = ref<HTMLElement | null>(null)
const presenterFlashOverlay = ref<HTMLElement | null>(null)
const continuingMediaHost = ref<HTMLElement | null>(null)
const pointerStartX = ref<number | null>(null)
const suppressClick = ref(false)
const slideNumberInput = ref('')
const contextMenu = ref<{ x: number; y: number } | null>(null)
type PresentationTool = 'pointer' | 'pen' | 'highlighter' | 'laser'
interface InkStroke { tool: 'pen' | 'highlighter'; points: string }
const activeTool = ref<PresentationTool>('pointer')
const inkBySlide = ref<Record<string, InkStroke[]>>({})
const draftStroke = ref<InkStroke | null>(null)
const laserPoint = ref<{ x: number; y: number } | null>(null)
const zoom = ref(1)
const zoomPan = ref({ x: 0, y: 0 })
let zoomBeforeReading: number | undefined
let zoomPanBeforeReading: { x: number; y: number } | undefined
let slideNumberTimer: ReturnType<typeof setTimeout> | undefined
let slideAdvanceTimer: ReturnType<typeof setTimeout> | undefined
let buildAutoAdvanceTimer: ReturnType<typeof setTimeout> | undefined
let presenterTimer: ReturnType<typeof setInterval> | undefined
let projectorChannel: BroadcastChannel | undefined
let projectorWindow: Window | null = null
let projectorChannelId = ''
let lastWheelAt = 0
const touchPoints = new Map<number, { x: number; y: number }>()
const continuingMedia = new Map<HTMLMediaElement, { endIndex: number; timer?: ReturnType<typeof setTimeout>; onEnded: () => void; frame?: HTMLElement }>()
let continuingMediaResizeObserver: ResizeObserver | undefined
let pinchStartDistance = 0
let pinchStartZoom = 1
let wasPinching = false
let panStart: { pointerId: number; x: number; y: number; offsetX: number; offsetY: number; moved: boolean } | undefined
const currentSlide = computed(() => deck.value?.slides[currentIndex.value] || null)
const renderingWarnings = computed(() => [...(deck.value?.warnings || []), ...(transitionSoundWarning.value ? [transitionSoundWarning.value] : [])])
const shownSlideIndices = computed(() => (deck.value?.slides || []).reduce<number[]>((indices, slide, index) => {
  if (!slide.hidden) indices.push(index)
  return indices
}, []))
const firstShownSlideIndex = computed(() => shownSlideIndices.value[0] ?? 0)
const lastShownSlideIndex = computed(() => shownSlideIndices.value[shownSlideIndices.value.length - 1] ?? Math.max(0, (deck.value?.slides.length || 1) - 1))
const nextShownSlideIndex = computed(() => shownSlideIndices.value.find(index => index > currentIndex.value)
  ?? (deck.value?.showSettings?.loop ? shownSlideIndices.value[0] : undefined))
const previousShownSlideIndex = computed(() => {
  const indices = shownSlideIndices.value
  for (let offset = indices.length - 1; offset >= 0; offset--) {
    if (indices[offset]! < currentIndex.value) return indices[offset]
  }
  return undefined
})
const thumbnailSlides = computed(() => deck.value?.slides.slice(thumbnailStart.value, thumbnailEnd.value).map((slide, offset) => ({ slide, index: thumbnailStart.value + offset })) || [])
const thumbnailCount = computed(() => deck.value?.slides.length || 0)
const thumbnailBeforeStyle = computed(() => thumbnailAxis.value === 'vertical'
  ? { height: `${Math.max(0, thumbnailStart.value * thumbnailStep.value - thumbnailGap.value)}px` }
  : { width: `${Math.max(0, thumbnailStart.value * thumbnailStep.value - thumbnailGap.value)}px` })
const thumbnailAfterStyle = computed(() => thumbnailAxis.value === 'vertical'
  ? { height: `${Math.max(0, (thumbnailCount.value - thumbnailEnd.value) * thumbnailStep.value - thumbnailGap.value)}px` }
  : { width: `${Math.max(0, (thumbnailCount.value - thumbnailEnd.value) * thumbnailStep.value - thumbnailGap.value)}px` })
const thumbnailOverscan = 3
const sorterPageSize = 48
const sorterSlides = computed(() => deck.value?.slides.slice(sorterPageStart.value, sorterPageStart.value + sorterPageSize).map((slide, offset) => ({ slide, index: sorterPageStart.value + offset })) || [])
const sorterPageEnd = computed(() => Math.min(sorterPageStart.value + sorterPageSize, deck.value?.slides.length || 0))
const currentNumber = computed(() => deck.value ? currentIndex.value + 1 : 0)
const showAnimations = computed(() => deck.value?.showSettings?.showAnimation !== false)
const showNarration = computed(() => deck.value?.showSettings?.showNarration === true)
const animationCount = computed(() => showAnimations.value ? currentSlide.value?.animationSteps?.length || 0 : 0)
const slideTimelineKey = computed(() => [
  deckGeneration.value,
  currentIndex.value,
  currentSlide.value?.id || '',
  animationStep.value,
  animationCount.value,
  transitionPlayback.value?.startedAtEpoch || '',
  triggerPlayback.value.map(({ sequence, step, startedAt }) => `${sequence}.${step}.${startedAt}`).join(',')
].join(':'))
const progress = computed(() => deck.value ? (currentNumber.value / deck.value.slides.length) * 100 : 0)
const currentInk = computed(() => currentSlide.value ? inkBySlide.value[currentSlide.value.id] || [] : [])
const presenterTime = computed(() => `${String(Math.floor(presenterElapsedSeconds.value / 60)).padStart(2, '0')}:${String(presenterElapsedSeconds.value % 60).padStart(2, '0')}`)

const randomTransitionEffects = ['fade', 'wipe', 'zoom', 'circle', 'diamond', 'plus', 'wedge', 'newsflash', 'honeycomb', 'blinds', 'checker', 'comb', 'dissolve', 'wheel', 'randomBar', 'strips', 'cover', 'wheelReverse', 'flash', 'push'] as const

function createTransitionPlayback(slide: PptxSlide | null): PptxTransitionPlayback | null {
  if (!slide?.transition) return null
  const source = slide.transition
  const transition = source.effect === 'random'
    ? { ...source, effect: randomTransitionEffects[Math.floor(Math.random() * randomTransitionEffects.length)]! }
    : source
  return { slideId: slide.id, startedAtEpoch: performance.timeOrigin + performance.now(), transition }
}

function stopTransitionSound(unload = false) {
  transitionSoundPlaybackId++
  transitionSoundShouldPlay = false
  const audio = transitionAudio.value
  if (!audio) return
  audio.loop = false
  audio.pause()
  if (audio.readyState > 0) audio.currentTime = 0
  if (unload) {
    audio.removeAttribute('src')
    audio.load()
    audio.pause()
  }
}

function playTransitionSound(playback: PptxTransitionPlayback | null) {
  if (!playback) return
  for (const action of playback.transition.soundActions || []) {
    stopTransitionSound(action.type === 'end')
    transitionSoundWarning.value = ''
    if (action.type === 'end') continue
    const audio = transitionAudio.value
    if (!audio) continue
    const playbackId = transitionSoundPlaybackId
    transitionSoundShouldPlay = true
    audio.src = action.url
    audio.loop = action.loop
    audio.currentTime = 0
    void audio.play().then(() => {
      if (playbackId !== transitionSoundPlaybackId && !transitionSoundShouldPlay) stopTransitionSound(true)
    }).catch(() => {
      if (playbackId === transitionSoundPlaybackId) {
        transitionSoundShouldPlay = false
        transitionSoundWarning.value = 'The browser could not play the transition sound.'
      }
    })
  }
}

watch(currentIndex, () => {
  transitionPlayback.value = createTransitionPlayback(currentSlide.value)
  void nextTick(ensureActiveThumbnailVisible)
}, { flush: 'sync' })
watch(transitionPlayback, playTransitionSound, { flush: 'sync' })
useSlideTransition({ slide: currentSlide, shell: stageShell, incomingHost: stageSlideHost, underlay: transitionUnderlay, flash: flashOverlay, playback: transitionPlayback })
useSlideTransition({ slide: currentSlide, shell: presenterStageShell, incomingHost: presenterSlideHost, underlay: presenterTransitionUnderlay, flash: presenterFlashOverlay, playback: transitionPlayback })

function stopContinuingMedia(media: HTMLMediaElement, pause = true) {
  const state = continuingMedia.get(media)
  if (state?.timer) clearTimeout(state.timer)
  if (state) media.removeEventListener('ended', state.onEnded)
  continuingMedia.delete(media)
  if (pause) media.pause()
  media.remove()
  state?.frame?.remove()
  if (!continuingMedia.size && continuingMediaHost.value) {
    continuingMediaHost.value.style.width = '0px'
    continuingMediaHost.value.style.height = '0px'
    continuingMediaHost.value.style.translate = ''
    continuingMediaHost.value.style.transform = ''
  }
}

function positionContinuingMediaHost() {
  const host = continuingMediaHost.value
  const slide = stageSlideHost.value
  if (!host || !slide || !continuingMedia.size) return
  host.style.left = `${slide.offsetLeft}px`
  host.style.top = `${slide.offsetTop}px`
  host.style.width = `${slide.offsetWidth}px`
  host.style.height = `${slide.offsetHeight}px`
  host.style.translate = slide.style.translate
  host.style.transform = slide.style.transform
  host.style.transformOrigin = getComputedStyle(slide).transformOrigin
}

function stopAllAudio() {
  for (const audio of document.querySelectorAll<HTMLAudioElement>('audio')) audio.pause()
  for (const media of [...continuingMedia.keys()]) stopContinuingMedia(media, false)
}

function continueMediaAcrossSlides(nextIndex: number, previousIndex: number) {
  for (const [media, state] of continuingMedia) {
    if (nextIndex < previousIndex || nextIndex >= state.endIndex || media.ended) stopContinuingMedia(media)
  }
  for (const media of stageShell.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || []) {
    if (continuingMedia.has(media)) continue
    const slideCount = Number(media.dataset.mediaSlideCount)
    if (media.paused || media.ended || !Number.isInteger(slideCount) || slideCount <= 1 || nextIndex <= previousIndex || nextIndex >= previousIndex + slideCount) {
      media.pause()
      continue
    }
    if (!(media instanceof HTMLAudioElement || media instanceof HTMLVideoElement) || !continuingMediaHost.value) {
      media.pause()
      continue
    }
    const continuation = media.cloneNode(false) as HTMLMediaElement
    continuation.src = media.currentSrc || media.src
    continuation.preload = 'auto'
    continuation.volume = media.volume
    continuation.muted = media.muted
    continuation.playbackRate = media.playbackRate
    continuation.loop = media.loop
    continuation.addEventListener('click', event => event.stopPropagation())
    continuation.addEventListener('pointerdown', event => event.stopPropagation())
    continuation.dataset.mediaId = media.dataset.mediaId
    if (media.dataset.mediaActionEndAt) continuation.dataset.mediaActionEndAt = media.dataset.mediaActionEndAt
    const startAt = media.currentTime
    const onEnded = () => stopContinuingMedia(continuation, false)
    let frame: HTMLElement | undefined
    if (continuation instanceof HTMLVideoElement) {
      const sourceFrame = media.closest<HTMLElement>('.slide-element')
      const slideFrame = stageSlideHost.value?.querySelector<HTMLElement>('.slide-frame')
      if (!sourceFrame || !slideFrame || !slideFrame.offsetWidth || !slideFrame.offsetHeight) {
        media.pause()
        continue
      }
      const sourceStyle = getComputedStyle(sourceFrame)
      frame = document.createElement('div')
      frame.className = 'continuing-video-frame'
      Object.assign(frame.style, {
        left: `${sourceFrame.offsetLeft / slideFrame.offsetWidth * 100}%`,
        top: `${sourceFrame.offsetTop / slideFrame.offsetHeight * 100}%`,
        width: `${sourceFrame.offsetWidth / slideFrame.offsetWidth * 100}%`,
        height: `${sourceFrame.offsetHeight / slideFrame.offsetHeight * 100}%`,
        transform: sourceStyle.transform,
        transformOrigin: sourceStyle.transformOrigin,
        opacity: sourceStyle.opacity,
        clipPath: sourceStyle.clipPath,
      })
      frame.append(continuation)
    }
    const state: { endIndex: number; timer?: ReturnType<typeof setTimeout>; onEnded: () => void; frame?: HTMLElement } = { endIndex: previousIndex + slideCount, onEnded }
    if (frame) state.frame = frame
    const endAt = Number(continuation.dataset.mediaActionEndAt)
    if (Number.isFinite(endAt)) state.timer = setTimeout(() => stopContinuingMedia(continuation), Math.max(0, endAt - performance.now()))
    continuation.addEventListener('ended', onEnded, { once: true })
    continuingMedia.set(continuation, state)
    if (frame) continuingMediaHost.value.append(frame)
    else continuingMediaHost.value.append(continuation)
    positionContinuingMediaHost()
    const start = () => {
      if (!continuingMedia.has(continuation)) return
      try { continuation.currentTime = startAt } catch {}
      void continuation.play().then(() => media.pause()).catch(() => stopContinuingMedia(continuation, false))
    }
    if (continuation.readyState >= 1) start()
    else continuation.addEventListener('loadedmetadata', start, { once: true })
  }
  void nextTick(positionContinuingMediaHost)
}

function togglePresenterView() {
  presenterView.value = !presenterView.value
  if (presenterView.value) {
    leaveReadingView()
  }
  if (presenterTimer) clearInterval(presenterTimer)
  presenterTimer = undefined
  if (!presenterView.value) {
    void nextTick(() => presenterTrigger.value?.focus())
    return
  }
  presenterElapsedSeconds.value = 0
  const startedAt = Date.now()
  presenterTimer = setInterval(() => { presenterElapsedSeconds.value = Math.floor((Date.now() - startedAt) / 1000) }, 1000)
  activeTool.value = 'pointer'
  void nextTick(() => presenterExit.value?.focus())
}

function leaveReadingView() {
  if (!readingView.value) return
  readingView.value = false
  readingNotesOpen.value = false
  if (zoomBeforeReading !== undefined) setZoom(zoomBeforeReading)
  if (zoomPanBeforeReading) setZoomPan(zoomPanBeforeReading.x, zoomPanBeforeReading.y)
  zoomBeforeReading = undefined
  zoomPanBeforeReading = undefined
}

function toggleReadingView() {
  if (readingView.value) {
    leaveReadingView()
    void nextTick(() => readingTrigger.value?.focus())
    return
  }
  zoomBeforeReading = zoom.value
  zoomPanBeforeReading = { ...zoomPan.value }
  setZoom(1)
  readingView.value = true
  readingNotesOpen.value = false
  activeTool.value = 'pointer'
  laserPoint.value = null
  contextMenu.value = null
  void nextTick(() => readingExit.value?.focus())
}

function postProjectorMessage(message: PresenterMessage) {
  try {
    projectorChannel?.postMessage(message)
  } catch {
    projectorError.value = 'projectorSendFailed'
  }
}

function sendProjectorDocument() {
  if (!deck.value) {
    postProjectorMessage({ type: 'clear' })
    return
  }
  postProjectorMessage({
    type: 'document', name: deck.value.name, width: deck.value.width, height: deck.value.height,
    totalSlides: deck.value.slides.length,
    showAnimations: showAnimations.value,
    showNarration: showNarration.value,
    embeddedFonts: deck.value.embeddedFonts,
    fontWarnings: deck.value.warnings.filter(warning => /embedded font/i.test(warning)),
  })
}

function sendProjectorSlide() {
  if (!currentSlide.value) return
  postProjectorMessage({
    type: 'slide', slide: currentSlide.value, number: currentNumber.value, animationStep: animationStep.value,
    blank: screenOverlay.value, triggerPlayback: triggerPlayback.value, transitionPlayback: transitionPlayback.value,
    ink: currentInk.value.map(stroke => ({ ...stroke })), draftStroke: draftStroke.value && { ...draftStroke.value },
    laserPoint: activeTool.value === 'laser' && laserPoint.value ? { ...laserPoint.value } : null,
  })
}

function sendProjectorState() {
  postProjectorMessage({ type: 'state', animationStep: animationStep.value, blank: screenOverlay.value, triggerPlayback: triggerPlayback.value })
}

function sendProjectorLaserPoint(point: { x: number; y: number } | null) {
  if (currentSlide.value) postProjectorMessage({ type: 'laser-point', slideId: currentSlide.value.id, point: point && { ...point } })
}

function sendProjectorSnapshot() {
  sendProjectorDocument()
  sendProjectorSlide()
}

function closeProjectorWindow() {
  if (projectorWindow && !projectorWindow.closed) projectorWindow.close()
  projectorWindow = null
  projectorOpen.value = false
  projectorConnected.value = false
  projectorAcknowledgement.value = null
}

function onProjectorMessage(event: MessageEvent<PresenterMessage>) {
  const message = event.data
  if (!message || typeof message !== 'object') return
  if (message.type === 'ready') {
    sendProjectorSnapshot()
  } else if (message.type === 'rendered') {
    if (message.slideId !== currentSlide.value?.id || message.number !== currentNumber.value || message.animationStep !== animationStep.value || message.blank !== screenOverlay.value || message.triggerCount !== triggerPlayback.value.length) return
    projectorConnected.value = true
    projectorAcknowledgement.value = { slideId: message.slideId, number: message.number, animationStep: message.animationStep, blank: message.blank, triggerCount: message.triggerCount }
  } else if (message.type === 'closed') {
    projectorWindow = null
    projectorOpen.value = false
    projectorConnected.value = false
    projectorAcknowledgement.value = null
  } else if (message.type === 'input') {
    if (message.action === 'next') nextSlide()
    else if (message.action === 'previous') previousSlide()
    else if (message.action === 'end') goToLastSlide()
    else if (message.action === 'go-to' && Number.isSafeInteger(message.slideNumber)
      && message.slideNumber > 0 && message.slideNumber <= (deck.value?.slides.length || 0)) goToSlide(message.slideNumber - 1)
    else if (message.action === 'blank-black') screenOverlay.value = 'black'
    else if (message.action === 'blank-white') screenOverlay.value = 'white'
    else if (message.action === 'blank-clear') screenOverlay.value = null
    else if (message.action === 'close') closeProjectorWindow()
  }
}

function openProjectorWindow() {
  if (!deck.value) return
  projectorError.value = ''
  if (!('BroadcastChannel' in window)) {
    projectorError.value = 'projectorNotSupported'
    return
  }
  if (!projectorChannel) {
    const bytes = crypto.getRandomValues(new Uint8Array(16))
    projectorChannelId = Array.from(bytes, byte => byte.toString(16).padStart(2, '0')).join('')
    projectorChannel = new BroadcastChannel(presenterChannelName(projectorChannelId))
    projectorChannel.addEventListener('message', onProjectorMessage)
  }
  if (!projectorWindow || projectorWindow.closed) {
    const url = new URL('/projector', window.location.href)
    url.searchParams.set('channel', projectorChannelId)
    projectorWindow = window.open(url.href, `deckline-projector-${projectorChannelId}`, 'popup,width=1280,height=720')
    if (!projectorWindow) {
      projectorOpen.value = false
      projectorError.value = 'projectorPopupBlocked'
      return
    }
    projectorOpen.value = true
    projectorConnected.value = false
  } else {
    projectorOpen.value = true
    projectorWindow.focus()
  }
  sendProjectorSnapshot()
}

watch(deck, () => {
  projectorAcknowledgement.value = null
  projectorConnected.value = false
  sendProjectorDocument()
  void nextTick(() => {
    if (thumbnailList.value) {
      thumbnailList.value.scrollTop = 0
      thumbnailList.value.scrollLeft = 0
    }
    thumbnailStart.value = 0
    thumbnailEnd.value = Math.min(thumbnailCount.value, 12)
    updateThumbnailWindow()
  })
}, { flush: 'post' })
watch([deck, currentIndex], () => {
  draftStroke.value = null
  laserPoint.value = null
  triggerPlayback.value = []
  sendProjectorSlide()
}, { flush: 'post' })
watch([animationStep, screenOverlay, triggerPlayback], sendProjectorState, { flush: 'post' })
watch(activeTool, tool => {
  if (tool !== 'laser') {
    laserPoint.value = null
    sendProjectorLaserPoint(null)
  }
})

function toggleSorterView() {
  sorterView.value = !sorterView.value
  if (sorterView.value) {
    leaveReadingView()
  }
  draggedSlideIndex.value = null
  if (!sorterView.value) {
    void nextTick(() => sorterTrigger.value?.focus())
    return
  }
  sorterPageStart.value = Math.floor(currentIndex.value / sorterPageSize) * sorterPageSize
  void nextTick(() => sorterExit.value?.focus())
}

function reorderSlide(from: number, to: number) {
  if (!deck.value || from < 0 || to < 0 || from >= deck.value.slides.length || to >= deck.value.slides.length || from === to) return
  const selectedId = currentSlide.value?.id
  const slides = [...deck.value.slides]
  const [moved] = slides.splice(from, 1)
  if (!moved) return
  slides.splice(to, 0, moved)
  deck.value = { ...deck.value, slides }
  currentIndex.value = Math.max(0, slides.findIndex(slide => slide.id === selectedId))
  sorterPageStart.value = Math.floor(currentIndex.value / sorterPageSize) * sorterPageSize
}

function startSlideDrag(event: DragEvent, index: number) {
  draggedSlideIndex.value = index
  event.dataTransfer?.setData('text/plain', String(index))
  if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move'
}

function dropSlide(event: DragEvent, index: number) {
  event.preventDefault()
  const from = draggedSlideIndex.value ?? Number(event.dataTransfer?.getData('text/plain'))
  if (Number.isInteger(from)) reorderSlide(from, index)
  draggedSlideIndex.value = null
}

function stepSorterPage(direction: -1 | 1) {
  if (!deck.value) return
  const lastPageStart = Math.floor((deck.value.slides.length - 1) / sorterPageSize) * sorterPageSize
  sorterPageStart.value = Math.min(lastPageStart, Math.max(0, sorterPageStart.value + direction * sorterPageSize))
}

function openPicker() {
  fileInput.value?.click()
}

function clearDocument() {
  deckGeneration.value++
  presenterView.value = false
  sorterView.value = false
  readingView.value = false
  readingNotesOpen.value = false
  zoomBeforeReading = undefined
  zoomPanBeforeReading = undefined
  draggedSlideIndex.value = null
  stageShell.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]').forEach(media => media.pause())
  stopTransitionSound(true)
  transitionSoundWarning.value = ''
  for (const media of continuingMedia.keys()) stopContinuingMedia(media)
  if (slideAdvanceTimer) clearTimeout(slideAdvanceTimer)
  slideAdvanceTimer = undefined
  if (buildAutoAdvanceTimer) clearTimeout(buildAutoAdvanceTimer)
  buildAutoAdvanceTimer = undefined
  if (presenterTimer) clearInterval(presenterTimer)
  presenterTimer = undefined
  stageShell.value?.querySelector<HTMLElement>(':scope > .slide-host')?.getAnimations().forEach(animation => animation.cancel())
  transitionUnderlay.value?.querySelectorAll<HTMLElement>('.slide-host').forEach(host => host.getAnimations().forEach(animation => animation.cancel()))
  stageShell.value?.classList.remove('is-prism-transition', 'is-flip-transition', 'is-switch-transition')
  flashOverlay.value?.getAnimations().forEach(animation => animation.cancel())
  transitionUnderlay.value?.replaceChildren()
  if (deck.value) disposePptxDocument(deck.value)
  deck.value = null
  currentIndex.value = 0
  animationStep.value = 0
  inkBySlide.value = {}
  draftStroke.value = null
  laserPoint.value = null
  transitionPlayback.value = null
  activeTool.value = 'pointer'
  setZoom(1)
  panStart = undefined
  touchPoints.clear()
}

async function loadFile(file?: File) {
  if (!file) return
  const generation = ++fileLoadGeneration
  errorMessage.value = ''
  if (isLoading.value) stopPptxParserWorker()
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    errorMessage.value = 'Please choose a .pptx file. Legacy .ppt files are not supported.'
    isLoading.value = false
    if (fileInput.value) fileInput.value.value = ''
    return
  }
  isLoading.value = true
  try {
    const nextDocument = await parsePptxInWorker(file)
    if (generation !== fileLoadGeneration) {
      disposePptxDocument(nextDocument)
      return
    }
    clearDocument()
    deck.value = nextDocument
    currentIndex.value = firstShownSlideIndex.value
  } catch (error) {
    if (generation === fileLoadGeneration) errorMessage.value = error instanceof Error ? error.message : ''
  } finally {
    if (generation === fileLoadGeneration) {
      isLoading.value = false
      if (fileInput.value) fileInput.value.value = ''
    }
  }
}

function clearSlideNumberInput() {
  slideNumberInput.value = ''
  clearTimeout(slideNumberTimer)
  slideNumberTimer = undefined
}

function previousSlide() {
  clearSlideNumberInput()
  if (animationStep.value > 0) {
    const slide = currentSlide.value
    const stepSequence = slide?.animationStepSequences?.[animationStep.value - 1]
    if (stepSequence?.rewindsOnPrevious === false || (slide?.animationSequence?.rewindsOnPrevious === false && stepSequence?.rewindsOnPrevious !== true)) return
    let targetStep = stepSequence?.previousStepIndex !== undefined
      ? stepSequence.previousStepIndex + 1
      : animationStep.value - 1
    if (slide && slide.animationSequence?.previousAction === 'skipTimed' && deck.value?.showSettings?.useTimings !== false) {
      // bldP@advAuto starts the next build step by time; skip such intermediate steps while rewinding.
      // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.buildparagraph?view=openxml-3.0.1
      while (targetStep > 1) {
        const precedingActions = slide.animationSteps?.[targetStep - 2] || []
        const delays = new Set(precedingActions.flatMap(action => action.buildAutoAdvanceAfterMs === undefined ? [] : [action.buildAutoAdvanceAfterMs]))
        if (delays.size !== 1) break
        targetStep--
      }
    }
    if (targetStep >= 0) {
      animationStep.value = targetStep
    } else if (deck.value && previousShownSlideIndex.value !== undefined) {
      currentIndex.value = previousShownSlideIndex.value
      animationStep.value = animationCount.value
    }
  }
  else if (deck.value && previousShownSlideIndex.value !== undefined) {
    currentIndex.value = previousShownSlideIndex.value
    animationStep.value = animationCount.value
  }
}

function nextSlide() {
  clearSlideNumberInput()
  if (animationStep.value < animationCount.value) {
    const stepSequence = currentSlide.value?.animationStepSequences?.[animationStep.value - 1]
    if (stepSequence?.advancesOnNext === false || (currentSlide.value?.animationSequence?.advancesOnNext === false && stepSequence?.advancesOnNext !== true)) return
    animationStep.value++
  }
  else if (deck.value && nextShownSlideIndex.value !== undefined) {
    const nextIndex = nextShownSlideIndex.value
    const repeatedSlide = currentIndex.value === nextIndex
    currentIndex.value = nextIndex
    animationStep.value = 0
    if (repeatedSlide) scheduleSlideAdvance()
  }
}

function goToSlide(index: number) {
  clearSlideNumberInput()
  currentIndex.value = index
  animationStep.value = 0
}

function goToFirstSlide() {
  clearSlideNumberInput()
  currentIndex.value = firstShownSlideIndex.value
  animationStep.value = 0
}

function goToLastSlide() {
  clearSlideNumberInput()
  const slides = deck.value?.slides
  if (!slides?.length) return
  const index = lastShownSlideIndex.value
  currentIndex.value = index
  animationStep.value = animationCount.value
}

function activateHyperlink(hyperlink: PptxHyperlink) {
  if (!deck.value) return
  if (hyperlink.kind === 'external') {
    try {
      const url = new URL(hyperlink.url)
      if (['http:', 'https:', 'mailto:'].includes(url.protocol) && !url.username && !url.password) {
        window.open(url.href, '_blank', 'noopener,noreferrer')
      }
    } catch {
      // Invalid URLs are ignored even if a model is constructed outside the PPTX parser.
    }
    return
  }

  let targetIndex: number
  if (hyperlink.kind === 'slide') {
    targetIndex = deck.value.slides.findIndex(slide => slide.packagePath === hyperlink.targetPath)
  } else {
    targetIndex = hyperlink.jump === 'firstslide' ? firstShownSlideIndex.value
      : hyperlink.jump === 'lastslide' ? lastShownSlideIndex.value
        : hyperlink.jump === 'nextslide' ? nextShownSlideIndex.value ?? -1
          : hyperlink.jump === 'previousslide' ? previousShownSlideIndex.value ?? -1
            : (hyperlink.slideNumber || 0) - 1
  }
  if (targetIndex >= 0 && targetIndex < deck.value.slides.length) goToSlide(targetIndex)
}

function updateThumbnailWindow() {
  const list = thumbnailList.value
  const total = thumbnailCount.value
  if (!list || !deck.value) {
    thumbnailStart.value = 0
    thumbnailEnd.value = 0
    return
  }

  const axis = window.matchMedia('(max-width: 860px)').matches ? 'horizontal' : 'vertical'
  const style = getComputedStyle(list)
  const gap = Number.parseFloat(axis === 'horizontal' ? style.columnGap : style.rowGap) || 0
  const item = list.querySelector<HTMLElement>('.thumb-button')
  const size = item?.getBoundingClientRect()[axis === 'horizontal' ? 'width' : 'height']
  const step = size ? size + gap : thumbnailStep.value
  thumbnailAxis.value = axis
  thumbnailGap.value = gap
  if (step > gap) thumbnailStep.value = step

  const offset = axis === 'horizontal' ? list.scrollLeft : list.scrollTop
  const viewport = axis === 'horizontal' ? list.clientWidth : list.clientHeight
  const start = Math.max(0, Math.floor(offset / thumbnailStep.value) - thumbnailOverscan)
  const end = Math.min(total, Math.max(start + 1, Math.ceil((offset + viewport) / thumbnailStep.value) + thumbnailOverscan))
  if (start !== thumbnailStart.value) thumbnailStart.value = start
  if (end !== thumbnailEnd.value) thumbnailEnd.value = end
}

function scheduleThumbnailUpdate() {
  if (thumbnailFrame) return
  thumbnailFrame = requestAnimationFrame(() => {
    thumbnailFrame = 0
    updateThumbnailWindow()
  })
}

function measureThumbnailItem(value: unknown) {
  if (value instanceof HTMLElement) updateThumbnailWindow()
}

function ensureActiveThumbnailVisible() {
  const list = thumbnailList.value
  const total = thumbnailCount.value
  if (!list || !total) return

  const hadFocus = list.contains(document.activeElement)
  const horizontal = window.matchMedia('(max-width: 860px)').matches
  const style = getComputedStyle(list)
  const paddingStart = Number.parseFloat(horizontal ? style.paddingLeft : style.paddingTop) || 0
  const paddingEnd = Number.parseFloat(horizontal ? style.paddingRight : style.paddingBottom) || 0
  const viewport = horizontal ? list.clientWidth : list.clientHeight
  const offset = horizontal ? list.scrollLeft : list.scrollTop
  const index = Math.min(total - 1, Math.max(0, currentIndex.value))
  const itemStart = paddingStart + index * thumbnailStep.value
  const itemEnd = itemStart + thumbnailStep.value - thumbnailGap.value
  let nextOffset = offset

  if (itemStart < offset + paddingStart) nextOffset = index * thumbnailStep.value
  else if (itemEnd > offset + viewport - paddingEnd) nextOffset = itemEnd - viewport + paddingEnd

  if (horizontal) list.scrollLeft = Math.max(0, nextOffset)
  else list.scrollTop = Math.max(0, nextOffset)
  updateThumbnailWindow()

  if (hadFocus) void nextTick(() => list.querySelector<HTMLButtonElement>(`[data-thumbnail-index="${index}"]`)?.focus({ preventScroll: true }))
}

function onKeydown(event: KeyboardEvent) {
  if (!deck.value) return
  const target = event.target instanceof HTMLElement ? event.target : null
  if (target && (target.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT', 'AUDIO', 'VIDEO'].includes(target.tagName) || ([' ', 'Enter'].includes(event.key) && !(event.key === 'Enter' && slideNumberInput.value) && Boolean(target.closest('button, a, [role="button"]'))))) return
  if (event.repeat && (event.key === 'Escape' || event.key === 'F5' || ['f', 's', 'g', 'b', '.', 'w', ','].includes(event.key.toLowerCase()))) return
  if (event.key === 'F5' && event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
    event.preventDefault()
    togglePresenterView()
    return
  }
  const sorterArrow = sorterView.value && event.altKey && !event.ctrlKey && !event.metaKey && ['ArrowLeft', 'ArrowUp', 'ArrowRight', 'ArrowDown'].includes(event.key)
  if ((event.altKey || event.ctrlKey || event.metaKey) && !sorterArrow) return
  if (slideNumberInput.value && !/^\d$/.test(event.key) && event.key !== 'Enter') clearSlideNumberInput()
  if (/^\d$/.test(event.key)) {
    event.preventDefault()
    const maxDigits = Math.max(3, String(deck.value.slides.length).length)
    slideNumberInput.value = `${slideNumberInput.value}${event.key}`.slice(-maxDigits)
    clearTimeout(slideNumberTimer)
    slideNumberTimer = setTimeout(clearSlideNumberInput, 2_000)
  } else if (event.key === 'Enter' && slideNumberInput.value) {
    event.preventDefault()
    const index = Number(slideNumberInput.value) - 1
    clearSlideNumberInput()
    if (index >= 0 && index < deck.value.slides.length) goToSlide(index)
  } else if (event.key === 'Escape') {
    if (presenterView.value) {
      togglePresenterView()
      return
    }
    if (sorterView.value) {
      toggleSorterView()
      return
    }
    if (readingView.value) {
      if (document.fullscreenElement) void document.exitFullscreen()
      else toggleReadingView()
      return
    }
    if (activeTool.value !== 'pointer') {
      activeTool.value = 'pointer'
      laserPoint.value = null
      return
    }
    contextMenu.value = null
    screenOverlay.value = null
    clearSlideNumberInput()
    if (document.fullscreenElement) void document.exitFullscreen()
  } else if (sorterView.value && event.altKey && ['ArrowLeft', 'ArrowUp', 'ArrowRight', 'ArrowDown'].includes(event.key)) {
    event.preventDefault()
    const delta = event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1
    reorderSlide(currentIndex.value, currentIndex.value + delta)
  } else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(event.key)) {
    event.preventDefault(); nextSlide()
  } else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key)) {
    event.preventDefault(); previousSlide()
  } else if (event.key === 'Home') {
    event.preventDefault()
    goToFirstSlide()
  } else if (event.key === 'End') {
    event.preventDefault()
    goToLastSlide()
  } else if (event.key === 'F5') {
    event.preventDefault()
    if (!event.shiftKey) {
      goToFirstSlide()
    }
    if (!document.fullscreenElement && stageShell.value) void stageShell.value.requestFullscreen().catch(() => {})
  } else if (event.key.toLowerCase() === 'f') {
    void toggleFullscreen()
  } else if (event.key.toLowerCase() === 's') {
    togglePresenterView()
  } else if (event.key.toLowerCase() === 'g') {
    toggleSorterView()
  } else if (['b', '.'].includes(event.key.toLowerCase())) {
    screenOverlay.value = screenOverlay.value === 'black' ? null : 'black'
  } else if (['w', ','].includes(event.key.toLowerCase())) {
    screenOverlay.value = screenOverlay.value === 'white' ? null : 'white'
  }
}

async function toggleFullscreen() {
  if (!stageShell.value) return
  if (document.fullscreenElement) await document.exitFullscreen()
  else await stageShell.value.requestFullscreen()
}

function onDrop(event: DragEvent) {
  isDragging.value = false
  void loadFile(event.dataTransfer?.files[0])
}

function onPointerDown(event: PointerEvent) {
  if (event.pointerType !== 'touch') return
  touchPoints.set(event.pointerId, { x: event.clientX, y: event.clientY })
  try { (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId) } catch {}
  if (touchPoints.size === 1 && zoom.value > 1) {
    pointerStartX.value = null
    panStart = { pointerId: event.pointerId, x: event.clientX, y: event.clientY, offsetX: zoomPan.value.x, offsetY: zoomPan.value.y, moved: false }
  } else if (touchPoints.size === 1) pointerStartX.value = event.clientX
  else if (touchPoints.size === 2) {
    panStart = undefined
    pointerStartX.value = null
    const [first, second] = [...touchPoints.values()]
    pinchStartDistance = Math.hypot(second!.x - first!.x, second!.y - first!.y)
    pinchStartZoom = zoom.value
    wasPinching = true
  }
}

function onPointerMove(event: PointerEvent) {
  if (event.pointerType !== 'touch' || !touchPoints.has(event.pointerId)) return
  const previousTouch = touchPoints.get(event.pointerId)!
  touchPoints.set(event.pointerId, { x: event.clientX, y: event.clientY })
  if (touchPoints.size === 1 && panStart?.pointerId === event.pointerId) {
    const deltaX = event.clientX - panStart.x
    const deltaY = event.clientY - panStart.y
    if (Math.hypot(deltaX, deltaY) >= 5) panStart.moved = true
    if (panStart.moved) setZoomPan(panStart.offsetX + deltaX, panStart.offsetY + deltaY)
    return
  }
  if (touchPoints.size < 2 || !pinchStartDistance) return
  const [firstEntry, secondEntry] = [...touchPoints.entries()]
  if (!firstEntry || !secondEntry) return
  const first = firstEntry[1]
  const second = secondEntry[1]
  const other = event.pointerId === firstEntry[0] ? second : first
  const previousCenterX = (previousTouch.x + other.x) / 2
  const previousCenterY = (previousTouch.y + other.y) / 2
  const centerX = (first.x + second.x) / 2
  const centerY = (first.y + second.y) / 2
  setZoomAt(pinchStartZoom * Math.hypot(second.x - first.x, second.y - first.y) / pinchStartDistance, previousCenterX, previousCenterY, centerX, centerY)
}

function onPointerUp(event: PointerEvent, allowSwipe = true) {
  if (event.pointerType !== 'touch') return
  touchPoints.delete(event.pointerId)
  if (panStart?.pointerId === event.pointerId) {
    if (panStart.moved) {
      suppressClick.value = true
      window.setTimeout(() => { suppressClick.value = false }, 400)
    }
    panStart = undefined
    pointerStartX.value = null
    return
  }
  if (wasPinching) {
    if (touchPoints.size < 2) {
      wasPinching = false
      pinchStartDistance = 0
      pointerStartX.value = null
      suppressClick.value = true
      window.setTimeout(() => { suppressClick.value = false }, 400)
    }
    return
  }
  if (!allowSwipe) {
    pointerStartX.value = null
    return
  }
  if (pointerStartX.value === null) return
  const distance = event.clientX - pointerStartX.value
  pointerStartX.value = null
  if (Math.abs(distance) < 48) return
  suppressClick.value = true
  if (distance < 0) nextSlide()
  else previousSlide()
  window.setTimeout(() => { suppressClick.value = false }, 400)
}

function onPointerCancel(event: PointerEvent) {
  onPointerUp(event, false)
}

function onWheel(event: WheelEvent) {
  if (!deck.value || Math.abs(event.deltaY) < 12) return
  if (event.ctrlKey) {
    setZoomAt(zoom.value + (event.deltaY < 0 ? .1 : -.1), event.clientX, event.clientY, event.clientX, event.clientY)
    return
  }
  if (performance.now() - lastWheelAt < 350) return
  lastWheelAt = performance.now()
  contextMenu.value = null
  if (event.deltaY > 0) nextSlide()
  else previousSlide()
}

function setZoom(value: number) {
  const bounds = stageShell.value?.getBoundingClientRect()
  const centerX = bounds ? bounds.left + bounds.width / 2 : 0
  const centerY = bounds ? bounds.top + bounds.height / 2 : 0
  setZoomAt(value, centerX, centerY, centerX, centerY)
}

function setZoomAt(value: number, fromX: number, fromY: number, toX: number, toY: number) {
  const previousZoom = zoom.value
  const nextZoom = Math.min(3, Math.max(.5, Math.round(value * 100) / 100))
  const bounds = stageShell.value?.getBoundingClientRect()
  const centerX = bounds ? bounds.left + bounds.width / 2 : 0
  const centerY = bounds ? bounds.top + bounds.height / 2 : 0
  const scale = nextZoom / previousZoom
  const panX = toX - centerX - scale * (fromX - centerX - zoomPan.value.x)
  const panY = toY - centerY - scale * (fromY - centerY - zoomPan.value.y)
  zoom.value = nextZoom
  setZoomPan(panX, panY)
}

function setZoomPan(x: number, y: number) {
  const host = stageSlideHost.value
  if (zoom.value <= 1 || !host) {
    zoomPan.value = { x: 0, y: 0 }
    return
  }
  const maxX = host.offsetWidth * (zoom.value - 1) / 2
  const maxY = host.offsetHeight * (zoom.value - 1) / 2
  zoomPan.value = {
    x: Math.max(-maxX, Math.min(maxX, x)),
    y: Math.max(-maxY, Math.min(maxY, y)),
  }
}

function inkPoint(event: PointerEvent) {
  const slide = (event.currentTarget as SVGElement).getBoundingClientRect()
  return {
    x: (event.clientX - slide.left) / slide.width * (deck.value?.width || 1),
    y: (event.clientY - slide.top) / slide.height * (deck.value?.height || 1),
  }
}

function onInkPointerDown(event: PointerEvent) {
  if (activeTool.value === 'pointer') return
  const point = inkPoint(event)
  try { (event.currentTarget as SVGElement).setPointerCapture(event.pointerId) } catch {}
  if (activeTool.value === 'laser') {
    laserPoint.value = point
    sendProjectorLaserPoint(point)
  } else if (currentSlide.value) {
    draftStroke.value = { tool: activeTool.value, points: `${point.x},${point.y}` }
    postProjectorMessage({ type: 'ink-start', slideId: currentSlide.value.id, stroke: { ...draftStroke.value } })
  }
}

function onInkPointerMove(event: PointerEvent) {
  if (activeTool.value !== 'laser' && !event.buttons && event.pointerType !== 'touch') return
  const point = inkPoint(event)
  if (activeTool.value === 'laser') {
    laserPoint.value = point
    sendProjectorLaserPoint(point)
  } else if (draftStroke.value) {
    draftStroke.value = { ...draftStroke.value, points: `${draftStroke.value.points} ${point.x},${point.y}` }
    if (currentSlide.value) postProjectorMessage({ type: 'ink-point', slideId: currentSlide.value.id, point })
  }
}

function onInkPointerUp() {
  const stroke = draftStroke.value
  if (stroke && currentSlide.value && stroke.points.includes(' ')) {
    inkBySlide.value = { ...inkBySlide.value, [currentSlide.value.id]: [...currentInk.value, stroke] }
  }
  if (stroke && currentSlide.value) postProjectorMessage({ type: 'ink-end', slideId: currentSlide.value.id })
  draftStroke.value = null
  if (laserPoint.value) sendProjectorLaserPoint(null)
  laserPoint.value = null
}

function onInkPointerLeave() {
  if (activeTool.value !== 'laser') return
  laserPoint.value = null
  sendProjectorLaserPoint(null)
}

function clearInk() {
  if (!currentSlide.value) return
  inkBySlide.value = { ...inkBySlide.value, [currentSlide.value.id]: [] }
  draftStroke.value = null
  postProjectorMessage({ type: 'ink-clear', slideId: currentSlide.value.id })
}

function openContextMenu(event: MouseEvent) {
  if (!deck.value || !stageShell.value) return
  const bounds = stageShell.value.getBoundingClientRect()
  contextMenu.value = {
    x: Math.max(8, Math.min(event.clientX - bounds.left, bounds.width - 170)),
    y: Math.max(8, Math.min(event.clientY - bounds.top, bounds.height - 190)),
  }
}

function activateHyperlinkFromElement(target: HTMLElement) {
  try {
    activateHyperlink(JSON.parse(target.dataset.pptxHyperlink || '') as PptxHyperlink)
  } catch {
    // Invalid link metadata is ignored without advancing the presentation.
  }
}

function onStageClick(event: MouseEvent) {
  const eventTarget = event.target instanceof Element ? event.target : undefined
  const hyperlinkTarget = eventTarget?.closest<HTMLElement>('[data-pptx-hyperlink]')
  if (hyperlinkTarget && !eventTarget?.closest('audio, video, button')) {
    event.preventDefault()
    event.stopPropagation()
    activateHyperlinkFromElement(hyperlinkTarget)
    return
  }
  if (contextMenu.value) {
    contextMenu.value = null
    return
  }
  if (suppressClick.value) {
    suppressClick.value = false
    return
  }
  if (screenOverlay.value) screenOverlay.value = null
  else if (deck.value && currentSlide.value && (animationStep.value < animationCount.value || currentSlide.value.advanceOnClick !== false)) nextSlide()
}

function onStageHyperlinkKeyActivate(event: KeyboardEvent) {
  const hyperlinkTarget = event.target instanceof Element
    ? event.target.closest<HTMLElement>('.slide-element[role="link"][data-pptx-hyperlink]')
    : null
  if (!hyperlinkTarget) return
  event.preventDefault()
  event.stopPropagation()
  activateHyperlinkFromElement(hyperlinkTarget)
}

function scheduleSlideAdvance() {
  if (slideAdvanceTimer) clearTimeout(slideAdvanceTimer)
  slideAdvanceTimer = undefined
  if (deck.value?.showSettings?.useTimings === false) return
  const slide = currentSlide.value
  const delayMs = slide?.advanceAfterMs
  const timelineKey = slideTimelineKey.value
  // PowerPoint starts its After timer when the slide's final animation/effect finishes.
  // https://support.microsoft.com/en-us/powerpoint/set-the-timing-and-speed-of-a-transition
  if (!slide || delayMs === undefined || nextShownSlideIndex.value === undefined
    || animationStep.value < animationCount.value
    || slideTimelineReadyKey.value !== timelineKey) return
  const slideIndex = currentIndex.value
  const now = Date.now()
  // OOXML advTm schedules when the slide transition starts; p14:dur controls the
  // transition's own duration and must not be added to the auto-advance delay.
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition?view=openxml-3.0.1
  const dueAt = now + delayMs
  const advance = () => {
    const remaining = dueAt - Date.now()
    if (remaining > 0) {
      slideAdvanceTimer = setTimeout(advance, Math.min(remaining, 2_147_483_647))
      return
    }
    if (currentIndex.value !== slideIndex || currentSlide.value !== slide
      || animationStep.value < animationCount.value || slideTimelineReadyKey.value !== timelineKey) return
    nextSlide()
  }
  slideAdvanceTimer = setTimeout(advance, Math.min(delayMs, 2_147_483_647))
}

function onSlideTimelineSettled(settled: boolean) {
  slideTimelineReadyKey.value = settled ? slideTimelineKey.value : ''
  scheduleSlideAdvance()
}

function scheduleBuildAutoAdvance() {
  if (buildAutoAdvanceTimer) clearTimeout(buildAutoAdvanceTimer)
  buildAutoAdvanceTimer = undefined
  if (deck.value?.showSettings?.useTimings === false) return
  const slide = currentSlide.value
  const stepIndex = animationStep.value - 1
  if (!slide || stepIndex < 0 || animationStep.value >= animationCount.value) return
  const delays = new Set((slide.animationSteps?.[stepIndex] || [])
    .flatMap(action => action.buildAutoAdvanceAfterMs === undefined ? [] : [action.buildAutoAdvanceAfterMs]))
  if (delays.size !== 1) return
  const delayMs = delays.values().next().value as number | undefined
  if (delayMs === undefined || !Number.isFinite(delayMs) || delayMs < 0) return
  const slideIndex = currentIndex.value
  const step = animationStep.value
  const dueAt = Date.now() + delayMs
  const advance = () => {
    const remaining = dueAt - Date.now()
    if (remaining > 0) {
      buildAutoAdvanceTimer = setTimeout(advance, Math.min(remaining, 2_147_483_647))
      return
    }
    if (currentIndex.value !== slideIndex || currentSlide.value !== slide || animationStep.value !== step || step >= animationCount.value) return
    animationStep.value++
  }
  buildAutoAdvanceTimer = setTimeout(advance, Math.min(delayMs, 2_147_483_647))
}

watch(currentIndex, (nextIndex, previousIndex) => {
  if (nextIndex === previousIndex) return
  continueMediaAcrossSlides(nextIndex, previousIndex)
  draftStroke.value = null
  laserPoint.value = null
  setZoom(1)
  void nextTick(positionContinuingMediaHost)
}, { flush: 'sync' })

watch([zoom, () => zoomPan.value.x, () => zoomPan.value.y], () => {
  void nextTick(positionContinuingMediaHost)
}, { flush: 'post' })

watch([() => deck.value, () => currentIndex.value, animationStep, animationCount, triggerPlayback, () => deck.value?.showSettings?.useTimings], scheduleSlideAdvance, { immediate: true })
watch([currentSlide, animationStep, animationCount, () => deck.value?.showSettings?.useTimings], scheduleBuildAutoAdvance, { flush: 'post' })

onMounted(() => {
  window.addEventListener('keydown', onKeydown)
  window.addEventListener('resize', scheduleThumbnailUpdate)
  window.addEventListener('resize', positionContinuingMediaHost)
  document.addEventListener('fullscreenchange', positionContinuingMediaHost)
  if (stageShell.value) {
    continuingMediaResizeObserver = new ResizeObserver(positionContinuingMediaHost)
    continuingMediaResizeObserver.observe(stageShell.value)
  }
  void nextTick(updateThumbnailWindow)
})
onBeforeUnmount(() => {
  fileLoadGeneration++
  window.removeEventListener('keydown', onKeydown)
  window.removeEventListener('resize', scheduleThumbnailUpdate)
  window.removeEventListener('resize', positionContinuingMediaHost)
  document.removeEventListener('fullscreenchange', positionContinuingMediaHost)
  continuingMediaResizeObserver?.disconnect()
  continuingMediaResizeObserver = undefined
  if (thumbnailFrame) cancelAnimationFrame(thumbnailFrame)
  clearTimeout(slideNumberTimer)
  if (presenterTimer) clearInterval(presenterTimer)
  if (projectorWindow && !projectorWindow.closed) projectorWindow.close()
  projectorWindow = null
  projectorChannel?.close()
  projectorChannel = undefined
  clearDocument()
  stopPptxParserWorker()
})

</script>

<template>
  <main class="app-shell" :class="{ 'is-reading-view': readingView }">
    <audio ref="transitionAudio" data-transition-sound hidden />
    <header class="topbar">
      <div class="brand" :aria-label="t('brandAria')">
        <span class="brand-mark">D</span>
        <span class="brand-name">deckline</span>
        <span class="brand-label">{{ t('brandLabel') }}</span>
      </div>
      <div class="topbar-actions">
        <div class="topbar-status"><span class="status-dot" />{{ t('localSession') }}</div>
        <button class="locale-toggle" :aria-label="t(locale === 'en' ? 'switchLanguageToChinese' : 'switchLanguageToEnglish')" @click="toggleLocale">{{ t(locale === 'en' ? 'switchToChinese' : 'switchToEnglish') }}</button>
      </div>
    </header>

    <div class="workspace" :inert="presenterView || sorterView">
      <aside class="deck-sidebar" :aria-label="t('slideList')">
        <div class="sidebar-head">
          <span class="sidebar-title">{{ t('slides') }}</span>
          <span class="slide-count">{{ deck ? String(deck.slides.length).padStart(2, '0') : '00' }}</span>
        </div>
        <div v-if="deck" ref="thumbnailList" class="thumb-list" @scroll="scheduleThumbnailUpdate">
          <div v-if="thumbnailStart > 0" class="thumb-virtual-spacer" :style="thumbnailBeforeStyle" aria-hidden="true" />
          <button
            v-for="{ slide, index } in thumbnailSlides"
            :key="slide.id"
            :ref="index === thumbnailStart ? measureThumbnailItem : undefined"
            class="thumb-button"
            :data-thumbnail-index="index"
            :class="{ 'is-active': currentIndex === index }"
            :aria-current="currentIndex === index ? 'page' : undefined"
            :aria-label="t(slide.hidden ? 'goToHiddenSlide' : 'goToSlide', { number: index + 1 })"
            @click="goToSlide(index)"
          >
            <div class="thumb-preview">
              <SlideStage :slide="slide" :width="deck.width" :height="deck.height" :number="index + 1" thumbnail />
            </div>
            <div class="thumb-meta"><span>{{ t(slide.hidden ? 'hidden' : 'slide') }}</span><span>{{ String(index + 1).padStart(2, '0') }}</span></div>
          </button>
          <div v-if="thumbnailEnd < thumbnailCount" class="thumb-virtual-spacer" :style="thumbnailAfterStyle" aria-hidden="true" />
        </div>
        <div v-else class="side-empty">{{ t('slideThumbnailsHint') }}</div>
        <div class="sidebar-footer">{{ t('poweredBy') }}</div>
      </aside>

      <section class="player-column" :aria-label="t('presentationPlayer')">
        <div class="deck-toolbar">
          <span class="file-icon">▤</span>
          <div class="deck-title-block">
            <div class="deck-title">{{ deck?.name || t('noPresentationOpen') }}</div>
            <div class="deck-subtitle">{{ deck ? t('slidesReady', { count: deck.slides.length }) : t('openPowerPointToBegin') }}</div>
          </div>
          <div class="tool-separator" />
          <button class="icon-button" :disabled="!deck || (animationStep === 0 && previousShownSlideIndex === undefined)" :aria-label="t('previousStep')" :title="t('previousStep')" @click="previousSlide">‹</button>
          <button class="icon-button" :disabled="!deck || (animationStep >= animationCount && nextShownSlideIndex === undefined)" :aria-label="t('nextStep')" :title="t('nextStep')" @click="nextSlide">›</button>
          <button class="text-button" :disabled="!deck" :title="t('fullscreenF')" @click="toggleFullscreen">⛶ <span>{{ t('present') }}</span></button>
          <button ref="sorterTrigger" class="tool-mode-button" data-tool="sorter" :disabled="!deck" :aria-label="t('openSlideSorter')" :title="t('slideSorterShortcut')" @click="toggleSorterView">{{ t('slideSorter') }}</button>
          <button ref="presenterTrigger" class="tool-mode-button" data-tool="speaker" :disabled="!deck" :aria-label="t('openSpeakerView')" :title="t('speakerViewShortcut')" @click="togglePresenterView">{{ t('speakerView') }}</button>
          <button ref="readingTrigger" class="tool-mode-button" data-tool="reading" :disabled="!deck" :aria-label="t('openReadingView')" :title="t('readingView')" @click="toggleReadingView">{{ t('reading') }}</button>
          <button v-for="tool in (['pointer', 'pen', 'highlighter', 'laser'] as const)" :key="tool" class="tool-mode-button" :class="{ 'is-active': activeTool === tool }" :data-tool="tool" :disabled="!deck" :aria-label="t(tool)" @click="activeTool = tool">{{ t(tool) }}</button>
          <button class="tool-mode-button" :disabled="!deck || !currentInk.length" data-tool="clear" :aria-label="t('clearAnnotations')" @click="clearInk">{{ t('clear') }}</button>
          <div class="tool-separator" />
          <button class="primary-button" :disabled="isLoading" @click="openPicker">{{ isLoading ? t('readingFile') : deck ? t('openAnother') : t('openPptx') }}</button>
          <input ref="fileInput" class="visually-hidden" type="file" tabindex="-1" aria-hidden="true" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation" @change="loadFile(($event.target as HTMLInputElement).files?.[0])">
        </div>

        <div
          ref="stageShell"
          class="stage-shell slide-transition-shell"
          @dragenter.prevent="isDragging = true"
          @dragover.prevent="isDragging = true"
          @dragleave.prevent="isDragging = false"
          @drop.prevent="onDrop"
          @pointerdown="onPointerDown"
          @pointermove="onPointerMove"
          @pointerup="onPointerUp"
          @pointercancel="onPointerCancel"
          @wheel.prevent="onWheel"
          @contextmenu.prevent="openContextMenu"
          @click="onStageClick"
          @keydown.enter="onStageHyperlinkKeyActivate"
        >
          <div class="stage-grid" />
          <div ref="transitionUnderlay" class="slide-transition-underlay" :class="{ 'is-pulling': currentSlide?.transition?.effect === 'pull', 'is-splitting-in': currentSlide?.transition?.effect === 'split' && currentSlide.transition.direction === 'in' }" aria-hidden="true" />
          <div ref="continuingMediaHost" class="continuing-media-host" />
          <div v-if="isDragging" class="upload-overlay">{{ t('dropPresentation') }}</div>
            <div v-if="currentSlide && deck" ref="stageSlideHost" class="slide-host" :data-slide-number="currentNumber" :data-transition-effect="currentSlide.transition?.effect" :data-transition-preset="currentSlide.transition?.presetName" :style="{ '--slide-ratio': String(deck.width / deck.height), translate: zoomPan.x + 'px ' + zoomPan.y + 'px', transform: 'scale(' + zoom + ')' }">
            <SlideStage :key="deckGeneration + '-' + currentSlide.id" :slide="currentSlide" :width="deck.width" :height="deck.height" :number="currentNumber" :animation-step="animationStep" :show-animations="showAnimations" :show-narration="showNarration" :use-timings="deck.showSettings?.useTimings !== false" @trigger-playback="triggerPlayback = $event" @timeline-settled="onSlideTimelineSettled" @activate-hyperlink="activateHyperlink" @stop-all-audio="stopAllAudio" />
            <svg
              class="ink-layer"
              :class="{ 'is-active': activeTool !== 'pointer' }"
              :viewBox="`0 0 ${deck.width} ${deck.height}`"
              preserveAspectRatio="none"
              :aria-label="t('annotationLayer')"
              @pointerdown.stop.prevent="onInkPointerDown"
              @pointermove.stop.prevent="onInkPointerMove"
              @pointerup.stop.prevent="onInkPointerUp"
              @pointercancel.stop.prevent="onInkPointerUp"
              @pointerleave="onInkPointerLeave"
            >
              <polyline v-for="(stroke, index) in currentInk" :key="index" :points="stroke.points" fill="none" :stroke="stroke.tool === 'pen' ? '#ef4444' : '#fde047'" :stroke-width="stroke.tool === 'pen' ? 3 : 16" :opacity="stroke.tool === 'pen' ? 1 : .4" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
              <polyline v-if="draftStroke" :points="draftStroke.points" fill="none" :stroke="draftStroke.tool === 'pen' ? '#ef4444' : '#fde047'" :stroke-width="draftStroke.tool === 'pen' ? 3 : 16" :opacity="draftStroke.tool === 'pen' ? 1 : .4" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
              <circle v-if="laserPoint" :cx="laserPoint.x" :cy="laserPoint.y" :r="deck.width / 120" fill="#ef4444" opacity=".9" />
            </svg>
            <div ref="flashOverlay" class="flash-transition-overlay" aria-hidden="true" />
          </div>
          <div v-else-if="!errorMessage" class="empty-state">
            <div class="empty-art"><span class="empty-art-back" /><span class="empty-art-front" /></div>
            <div class="empty-kicker">{{ t('quieterWay') }}</div>
            <h1>{{ t('deckInRoom') }}</h1>
            <p>{{ t('localPresentationDescription') }}</p>
            <button class="primary-button" :disabled="isLoading" @click.stop="openPicker">{{ isLoading ? t('readingPresentation') : t('choosePptx') }}</button>
            <div class="drop-hint">{{ t('dropFileAnywhere') }}</div>
          </div>
          <div v-if="screenOverlay" class="screen-overlay" :class="`is-${screenOverlay}`" @click.stop="screenOverlay = null" />
          <div v-if="slideNumberInput" class="slide-number-input" aria-live="polite">{{ slideNumberInput }}</div>
          <div v-if="contextMenu" class="presentation-menu" :style="{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }" role="menu" @click.stop>
            <button role="menuitem" :disabled="animationStep === 0 && previousShownSlideIndex === undefined" @click="previousSlide(); contextMenu = null">{{ t('previous') }}</button>
            <button role="menuitem" :disabled="animationStep >= animationCount && nextShownSlideIndex === undefined" @click="nextSlide(); contextMenu = null">{{ t('next') }}</button>
            <button role="menuitem" @click="screenOverlay = 'black'; contextMenu = null">{{ t('blackScreen') }}</button>
            <button role="menuitem" @click="screenOverlay = 'white'; contextMenu = null">{{ t('whiteScreen') }}</button>
          </div>
          <div v-if="errorMessage" class="error-card" role="alert">
            <strong>{{ t('cannotOpenPresentation') }}</strong>{{ translateError(errorMessage) }}
          </div>
          <details v-if="renderingWarnings.length" class="warning-panel" @click.stop>
            <summary>{{ t('renderingLimitations', { count: renderingWarnings.length }) }}</summary>
            <p v-if="locale === 'zh'" class="warning-detail-note">{{ t('warningDetails') }}</p>
            <ul class="warning-list">
              <li v-for="warning in renderingWarnings" :key="warning">
                <template v-if="locale === 'zh'">
                  <span>{{ translateWarning(warning) }}</span>
                  <details class="warning-original">
                    <summary>{{ t('warningOriginal') }}</summary>
                    <code>{{ warning }}</code>
                  </details>
                </template>
                <template v-else>{{ warning }}</template>
              </li>
            </ul>
          </details>
          <div v-if="deck" class="progress-track" aria-hidden="true"><div class="progress-value" :style="{ width: `${progress}%` }" /></div>
        </div>

        <div class="player-controls">
          <button class="icon-button" :disabled="!deck || (animationStep === 0 && previousShownSlideIndex === undefined)" :aria-label="t('previousStep')" @click="previousSlide">←</button>
          <div class="slide-position"><strong>{{ String(currentNumber).padStart(2, '0') }}</strong> / {{ String(deck?.slides.length || 0).padStart(2, '0') }}</div>
          <button class="icon-button" :disabled="!deck || (animationStep >= animationCount && nextShownSlideIndex === undefined)" :aria-label="t('nextStep')" @click="nextSlide">→</button>
          <button class="zoom-button" :disabled="!deck || zoom <= .5" :aria-label="t('zoomOut')" @click="setZoom(zoom - .25)">−</button>
          <span class="zoom-value">{{ Math.round(zoom * 100) }}%</span>
          <button class="zoom-button" :disabled="!deck || zoom >= 3" :aria-label="t('zoomIn')" @click="setZoom(zoom + .25)">+</button>
          <span class="shortcut-hint">{{ t('keyboardShortcuts') }}</span>
          <button v-if="readingView" class="reading-control-button" data-reading-notes :aria-expanded="readingNotesOpen" :aria-controls="readingNotesOpen ? 'reading-notes' : undefined" @click="readingNotesOpen = !readingNotesOpen">{{ readingNotesOpen ? t('hideNotes') : t('notes') }}</button>
          <button v-if="readingView" ref="readingExit" class="reading-control-button" data-reading-exit @click="toggleReadingView">{{ t('backToPlayer') }}</button>
          <button v-if="readingView" class="locale-toggle" :aria-label="t(locale === 'en' ? 'switchLanguageToChinese' : 'switchLanguageToEnglish')" @click="toggleLocale">{{ t(locale === 'en' ? 'switchToChinese' : 'switchToEnglish') }}</button>
        </div>
        <aside v-if="readingView && readingNotesOpen && currentSlide" id="reading-notes" class="reading-notes-panel" data-reading-notes-panel :aria-label="t('speakerNotes')">
          <header><strong>{{ t('speakerNotes') }}</strong><button :aria-label="t('hideSpeakerNotes')" @click="readingNotesOpen = false">×</button></header>
          <SpeakerNotes v-if="currentSlide.speakerNoteParagraphs?.length" :paragraphs="currentSlide.speakerNoteParagraphs" @activate-hyperlink="activateHyperlink" />
          <p v-else-if="currentSlide.speakerNotes" class="reading-notes-plain">{{ currentSlide.speakerNotes }}</p>
          <p v-else class="reading-notes-plain">{{ t('noSpeakerNotes') }}</p>
        </aside>
      </section>
    </div>
    <div v-if="presenterView && deck && currentSlide" class="presenter-view" data-presenter-view :data-current-slide="currentNumber" :data-animation-step="animationStep" role="dialog" aria-modal="true" :aria-label="t('speakerView')">
      <header class="presenter-header">
        <div><strong>{{ deck.name }}</strong><span> {{ t('speakerView').toUpperCase() }}</span></div>
        <div class="presenter-header-actions">
          <span class="presenter-clock" :aria-label="t('elapsedTime')">{{ presenterTime }}</span>
          <span class="presenter-display-status" :data-projector-connected="projectorConnected" :data-projector-number="projectorAcknowledgement?.number ?? 0" :data-projector-step="projectorAcknowledgement?.animationStep ?? -1" :data-projector-blank="projectorAcknowledgement?.blank || ''" :data-projector-trigger-count="projectorAcknowledgement?.triggerCount ?? -1">{{ projectorConnected ? t('secondDisplayConnected') : t('secondDisplayClosed') }}</span>
          <button class="text-button" data-open-projector @click="openProjectorWindow">{{ projectorOpen ? t('focusDisplay') : t('openSecondDisplay') }}</button>
          <span v-if="projectorError" class="projector-error" data-projector-error role="alert">{{ t(projectorError) }}</span>
          <button v-if="projectorOpen" class="text-button" data-close-projector @click="closeProjectorWindow">{{ t('closeDisplay') }}</button>
          <button class="locale-toggle" :aria-label="t(locale === 'en' ? 'switchLanguageToChinese' : 'switchLanguageToEnglish')" @click="toggleLocale">{{ t(locale === 'en' ? 'switchToChinese' : 'switchToEnglish') }}</button>
          <button ref="presenterExit" class="text-button" data-presenter-exit @click="togglePresenterView">{{ t('exit') }} <kbd>Esc</kbd></button>
        </div>
      </header>
      <div class="presenter-grid">
        <section class="presenter-current-panel" :aria-label="t('current') + ' ' + t('slides').toLowerCase()">
          <div class="presenter-panel-title"><span>{{ t('current') }}</span><span>{{ currentNumber }} / {{ deck.slides.length }}</span></div>
          <div class="presenter-current-frame">
            <div ref="presenterStageShell" class="slide-transition-shell presenter-transition-shell">
              <div ref="presenterTransitionUnderlay" class="slide-transition-underlay" :class="{ 'is-pulling': currentSlide?.transition?.effect === 'pull', 'is-splitting-in': currentSlide?.transition?.effect === 'split' && currentSlide.transition.direction === 'in' }" aria-hidden="true" />
              <div ref="presenterSlideHost" class="slide-host" :data-slide-number="currentNumber" :style="{ '--slide-ratio': String(deck.width / deck.height) }">
                <SlideStage class="presenter-large-slide" :style="{ '--slide-ratio': String(deck.width / deck.height) }" :slide="currentSlide" :width="deck.width" :height="deck.height" :number="currentNumber" :animation-step="animationStep" :show-animations="showAnimations" :show-narration="showNarration" :trigger-playback="triggerPlayback" thumbnail playback-preview />
              </div>
              <div ref="presenterFlashOverlay" class="flash-transition-overlay" aria-hidden="true" />
            </div>
          </div>
        </section>
        <aside class="presenter-side-panel">
          <section class="presenter-next-panel" :aria-label="t('next') + ' ' + t('slides').toLowerCase()">
            <div class="presenter-panel-title"><span>{{ t('next') }}</span><span>{{ nextShownSlideIndex !== undefined ? `${nextShownSlideIndex + 1} / ${deck.slides.length}` : t('endOfDeck') }}</span></div>
            <div v-if="nextShownSlideIndex !== undefined" class="presenter-next-frame">
              <SlideStage class="presenter-next-slide" :slide="deck.slides[nextShownSlideIndex]!" :width="deck.width" :height="deck.height" :number="nextShownSlideIndex + 1" thumbnail />
            </div>
            <div v-else class="presenter-end-card">{{ t('finalSlide') }}</div>
          </section>
          <section class="presenter-notes-panel" :aria-label="t('speakerNotes')">
            <div class="presenter-panel-title"><span>{{ t('speakerNotes') }}</span></div>
            <SpeakerNotes v-if="currentSlide.speakerNoteParagraphs?.length" :paragraphs="currentSlide.speakerNoteParagraphs" @activate-hyperlink="activateHyperlink" />
            <p v-else-if="currentSlide.speakerNotes" data-presenter-notes>{{ currentSlide.speakerNotes }}</p>
            <p v-else class="presenter-no-notes" data-presenter-notes>{{ t('noSpeakerNotes') }}</p>
          </section>
        </aside>
      </div>
      <footer class="presenter-footer">
        <button class="text-button" :disabled="animationStep === 0 && previousShownSlideIndex === undefined" @click="previousSlide">{{ t('previousSlide') }}</button>
        <span>{{ t('presenterKeyboardHelp', { time: presenterTime }) }}</span>
        <button class="text-button" :disabled="animationStep >= animationCount && nextShownSlideIndex === undefined" @click="nextSlide">{{ t('nextSlide') }}</button>
      </footer>
    </div>
    <div v-if="sorterView && deck" class="sorter-view" data-sorter-view role="dialog" aria-modal="true" :aria-label="t('slideSorter')">
      <header class="sorter-header">
        <div><strong>{{ t('slideSorter') }}</strong><span>{{ sorterPageStart + 1 }}–{{ sorterPageEnd }} / {{ deck.slides.length }}</span></div>
        <div class="presenter-header-actions">
          <button class="locale-toggle" :aria-label="t(locale === 'en' ? 'switchLanguageToChinese' : 'switchLanguageToEnglish')" @click="toggleLocale">{{ t(locale === 'en' ? 'switchToChinese' : 'switchToEnglish') }}</button>
          <button ref="sorterExit" class="text-button" data-sorter-exit @click="toggleSorterView">{{ t('close') }} <kbd>Esc</kbd></button>
        </div>
      </header>
      <div class="sorter-grid" role="list">
        <article
          v-for="{ slide, index } in sorterSlides"
          :key="slide.id"
          class="sorter-card"
          :class="{ 'is-active': currentIndex === index, 'is-dragging': draggedSlideIndex === index }"
          :data-slide-index="index"
          role="listitem"
          draggable="true"
          @dragstart="startSlideDrag($event, index)"
          @dragover.prevent
          @drop.stop="dropSlide($event, index)"
          @dragend="draggedSlideIndex = null"
        >
          <button class="sorter-slide-button" :aria-label="`${t(slide.hidden ? 'goToHiddenSlide' : 'goToSlide', { number: index + 1 })}: ${slide.name}`" @click="goToSlide(index); toggleSorterView()">
            <div class="sorter-preview" :style="{ aspectRatio: `${deck.width} / ${deck.height}` }"><SlideStage :slide="slide" :width="deck.width" :height="deck.height" :number="index + 1" thumbnail /></div>
            <div class="sorter-meta"><span>{{ String(index + 1).padStart(2, '0') }}</span><span>{{ slide.hidden ? t('hidden') + ' · ' : '' }}{{ slide.name }}</span></div>
          </button>
        </article>
      </div>
      <footer class="sorter-footer">
        <button class="text-button" :disabled="sorterPageStart === 0" @click="stepSorterPage(-1)">{{ t('previousGroup') }}</button>
        <span>{{ t('reorderSlidesHint') }}</span>
        <button class="text-button" :disabled="sorterPageEnd >= deck.slides.length" @click="stepSorterPage(1)">{{ t('nextGroup') }}</button>
      </footer>
    </div>
  </main>
</template>
