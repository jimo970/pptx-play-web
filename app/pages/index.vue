<script setup lang="ts">
import type { PptxDocument, PptxHyperlink, PptxSlide, PptxTransitionPlayback, PptxTriggerPlayback } from '~/utils/pptx'
import { parsePptxInWorker, stopPptxParserWorker } from '~/utils/pptx-parser-client'
import { presenterChannelName, type PresenterMessage } from '~/utils/presenter-channel'
import { useSlideTransition } from '~/composables/useSlideTransition'

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
const presenterElapsedSeconds = ref(0)
const projectorConnected = ref(false)
const projectorOpen = ref(false)
const projectorError = ref('')
const projectorAcknowledgement = ref<{ slideId: string; number: number; animationStep: number; blank: 'black' | 'white' | null; triggerCount: number } | null>(null)
const triggerPlayback = shallowRef<PptxTriggerPlayback[]>([])
const transitionPlayback = shallowRef<PptxTransitionPlayback | null>(null)
const presenterTrigger = ref<HTMLButtonElement | null>(null)
const presenterExit = ref<HTMLButtonElement | null>(null)
const sorterTrigger = ref<HTMLButtonElement | null>(null)
const sorterExit = ref<HTMLButtonElement | null>(null)
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
let slideNumberTimer: ReturnType<typeof setTimeout> | undefined
let slideAdvanceTimer: ReturnType<typeof setTimeout> | undefined
let presenterTimer: ReturnType<typeof setInterval> | undefined
let projectorChannel: BroadcastChannel | undefined
let projectorWindow: Window | null = null
let projectorChannelId = ''
let lastWheelAt = 0
const touchPoints = new Map<number, { x: number; y: number }>()
const continuingMedia = new Map<HTMLMediaElement, { endIndex: number; timer?: ReturnType<typeof setTimeout>; onEnded: () => void }>()
let pinchStartDistance = 0
let pinchStartZoom = 1
let wasPinching = false
const currentSlide = computed(() => deck.value?.slides[currentIndex.value] || null)
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
const animationCount = computed(() => currentSlide.value?.animationSteps?.length || 0)
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

watch(currentIndex, () => {
  transitionPlayback.value = createTransitionPlayback(currentSlide.value)
  void nextTick(ensureActiveThumbnailVisible)
}, { flush: 'sync' })
useSlideTransition({ slide: currentSlide, shell: stageShell, incomingHost: stageSlideHost, underlay: transitionUnderlay, flash: flashOverlay, playback: transitionPlayback })
useSlideTransition({ slide: currentSlide, shell: presenterStageShell, incomingHost: presenterSlideHost, underlay: presenterTransitionUnderlay, flash: presenterFlashOverlay, playback: transitionPlayback })

function stopContinuingMedia(media: HTMLMediaElement, pause = true) {
  const state = continuingMedia.get(media)
  if (state?.timer) clearTimeout(state.timer)
  if (state) media.removeEventListener('ended', state.onEnded)
  continuingMedia.delete(media)
  if (pause) media.pause()
  media.remove()
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
    if (!(media instanceof HTMLAudioElement) || !continuingMediaHost.value) {
      media.pause()
      continue
    }
    const continuation = new Audio(media.currentSrc || media.src)
    continuation.preload = 'auto'
    continuation.volume = media.volume
    continuation.muted = media.muted
    continuation.playbackRate = media.playbackRate
    continuation.dataset.mediaId = media.dataset.mediaId
    if (media.dataset.mediaActionEndAt) continuation.dataset.mediaActionEndAt = media.dataset.mediaActionEndAt
    const startAt = media.currentTime
    const onEnded = () => stopContinuingMedia(continuation, false)
    const state: { endIndex: number; timer?: ReturnType<typeof setTimeout>; onEnded: () => void } = { endIndex: previousIndex + slideCount, onEnded }
    const endAt = Number(continuation.dataset.mediaActionEndAt)
    if (Number.isFinite(endAt)) state.timer = setTimeout(() => stopContinuingMedia(continuation), Math.max(0, endAt - performance.now()))
    continuation.addEventListener('ended', onEnded, { once: true })
    continuingMedia.set(continuation, state)
    continuingMediaHost.value.append(continuation)
    const start = () => {
      if (!continuingMedia.has(continuation)) return
      try { continuation.currentTime = startAt } catch {}
      void continuation.play().then(() => media.pause()).catch(() => stopContinuingMedia(continuation, false))
    }
    if (continuation.readyState >= 1) start()
    else continuation.addEventListener('loadedmetadata', start, { once: true })
  }
}

function togglePresenterView() {
  presenterView.value = !presenterView.value
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

function postProjectorMessage(message: PresenterMessage) {
  try {
    projectorChannel?.postMessage(message)
  } catch {
    projectorError.value = 'The current slide could not be sent to the second display.'
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
  })
}

function sendProjectorSlide() {
  if (!currentSlide.value) return
  postProjectorMessage({
    type: 'slide', slide: currentSlide.value, number: currentNumber.value, animationStep: animationStep.value,
    blank: screenOverlay.value, triggerPlayback: triggerPlayback.value, transitionPlayback: transitionPlayback.value,
  })
}

function sendProjectorState() {
  postProjectorMessage({ type: 'state', animationStep: animationStep.value, blank: screenOverlay.value, triggerPlayback: triggerPlayback.value })
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
    projectorError.value = 'This browser does not support synchronized second displays.'
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
      projectorError.value = 'The browser blocked the projector window. Allow pop-ups for this site and try again.'
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
  triggerPlayback.value = []
  sendProjectorSlide()
}, { flush: 'post' })
watch([animationStep, screenOverlay, triggerPlayback], sendProjectorState, { flush: 'post' })

function toggleSorterView() {
  sorterView.value = !sorterView.value
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
  draggedSlideIndex.value = null
  stageShell.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]').forEach(media => media.pause())
  for (const media of continuingMedia.keys()) stopContinuingMedia(media)
  if (slideAdvanceTimer) clearTimeout(slideAdvanceTimer)
  slideAdvanceTimer = undefined
  if (presenterTimer) clearInterval(presenterTimer)
  presenterTimer = undefined
  stageShell.value?.querySelector<HTMLElement>(':scope > .slide-host')?.getAnimations().forEach(animation => animation.cancel())
  transitionUnderlay.value?.querySelectorAll<HTMLElement>('.slide-host').forEach(host => host.getAnimations().forEach(animation => animation.cancel()))
  stageShell.value?.classList.remove('is-prism-transition', 'is-flip-transition', 'is-switch-transition')
  flashOverlay.value?.getAnimations().forEach(animation => animation.cancel())
  transitionUnderlay.value?.replaceChildren()
  deck.value?.objectUrls.forEach(URL.revokeObjectURL)
  deck.value = null
  currentIndex.value = 0
  animationStep.value = 0
  inkBySlide.value = {}
  draftStroke.value = null
  laserPoint.value = null
  transitionPlayback.value = null
  activeTool.value = 'pointer'
  zoom.value = 1
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
      nextDocument.objectUrls.forEach(URL.revokeObjectURL)
      return
    }
    clearDocument()
    deck.value = nextDocument
    currentIndex.value = 0
  } catch (error) {
    if (generation === fileLoadGeneration) errorMessage.value = error instanceof Error ? error.message : 'The presentation could not be opened.'
  } finally {
    if (generation === fileLoadGeneration) {
      isLoading.value = false
      if (fileInput.value) fileInput.value.value = ''
    }
  }
}

function previousSlide() {
  if (animationStep.value > 0) {
    if (currentSlide.value?.animationSequence?.rewindsOnPrevious === false) return
    // withEffect/afterEffect groups are folded into their preceding click step, matching skipTimed's prior-step boundary.
    animationStep.value--
  }
  else if (deck.value && currentIndex.value > 0) {
    currentIndex.value--
    animationStep.value = deck.value.slides[currentIndex.value]?.animationSteps?.length || 0
  }
}

function nextSlide() {
  if (animationStep.value < animationCount.value) {
    if (currentSlide.value?.animationSequence?.advancesOnNext === false) return
    animationStep.value++
  }
  else if (deck.value && currentIndex.value < deck.value.slides.length - 1) {
    currentIndex.value++
    animationStep.value = 0
  }
}

function goToSlide(index: number) {
  currentIndex.value = index
  animationStep.value = 0
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
    targetIndex = hyperlink.jump === 'firstslide' ? 0
      : hyperlink.jump === 'lastslide' ? deck.value.slides.length - 1
        : hyperlink.jump === 'nextslide' ? currentIndex.value + 1
          : hyperlink.jump === 'previousslide' ? currentIndex.value - 1
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
  const target = event.target as HTMLElement | null
  if (target && ['INPUT', 'TEXTAREA', 'SELECT', 'AUDIO', 'VIDEO'].includes(target.tagName)) return
  if (/^\d$/.test(event.key) && !event.altKey && !event.ctrlKey && !event.metaKey) {
    event.preventDefault()
    slideNumberInput.value = `${slideNumberInput.value}${event.key}`.slice(-3)
    clearTimeout(slideNumberTimer)
    slideNumberTimer = setTimeout(() => { slideNumberInput.value = '' }, 2_000)
  } else if (event.key === 'Enter' && slideNumberInput.value) {
    event.preventDefault()
    const index = Number(slideNumberInput.value) - 1
    slideNumberInput.value = ''
    clearTimeout(slideNumberTimer)
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
    if (activeTool.value !== 'pointer') {
      activeTool.value = 'pointer'
      laserPoint.value = null
      return
    }
    contextMenu.value = null
    screenOverlay.value = null
    slideNumberInput.value = ''
    if (document.fullscreenElement) void document.exitFullscreen()
  } else if (sorterView.value && event.altKey && ['ArrowLeft', 'ArrowUp', 'ArrowRight', 'ArrowDown'].includes(event.key)) {
    event.preventDefault()
    const delta = event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1
    reorderSlide(currentIndex.value, currentIndex.value + delta)
  } else if (['ArrowRight', 'ArrowDown', 'PageDown', ' '].includes(event.key)) {
    event.preventDefault(); nextSlide()
  } else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key)) {
    event.preventDefault(); previousSlide()
  } else if (event.key === 'Home') {
    currentIndex.value = 0
    animationStep.value = 0
  } else if (event.key === 'End') {
    currentIndex.value = deck.value.slides.length - 1
    animationStep.value = deck.value.slides[currentIndex.value]?.animationSteps?.length || 0
  } else if (event.key === 'F5') {
    event.preventDefault()
    if (!event.shiftKey) {
      currentIndex.value = 0
      animationStep.value = 0
    }
    if (!document.fullscreenElement && stageShell.value) void stageShell.value.requestFullscreen().catch(() => {})
  } else if (event.key.toLowerCase() === 'f') {
    void toggleFullscreen()
  } else if (event.key.toLowerCase() === 's') {
    togglePresenterView()
  } else if (event.key.toLowerCase() === 'g') {
    toggleSorterView()
  } else if (event.key.toLowerCase() === 'b') {
    screenOverlay.value = screenOverlay.value === 'black' ? null : 'black'
  } else if (event.key.toLowerCase() === 'w') {
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
  if (touchPoints.size === 1) pointerStartX.value = event.clientX
  else if (touchPoints.size === 2) {
    const [first, second] = [...touchPoints.values()]
    pinchStartDistance = Math.hypot(second!.x - first!.x, second!.y - first!.y)
    pinchStartZoom = zoom.value
    wasPinching = true
  }
}

function onPointerMove(event: PointerEvent) {
  if (event.pointerType !== 'touch' || !touchPoints.has(event.pointerId)) return
  touchPoints.set(event.pointerId, { x: event.clientX, y: event.clientY })
  if (touchPoints.size < 2 || !pinchStartDistance) return
  const [first, second] = [...touchPoints.values()]
  setZoom(pinchStartZoom * Math.hypot(second!.x - first!.x, second!.y - first!.y) / pinchStartDistance)
}

function onPointerUp(event: PointerEvent, allowSwipe = true) {
  if (event.pointerType !== 'touch') return
  touchPoints.delete(event.pointerId)
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
    setZoom(zoom.value + (event.deltaY < 0 ? .1 : -.1))
    return
  }
  if (performance.now() - lastWheelAt < 350) return
  lastWheelAt = performance.now()
  contextMenu.value = null
  if (event.deltaY > 0) nextSlide()
  else previousSlide()
}

function setZoom(value: number) {
  zoom.value = Math.min(3, Math.max(.5, Math.round(value * 100) / 100))
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
  if (activeTool.value === 'laser') laserPoint.value = point
  else draftStroke.value = { tool: activeTool.value, points: `${point.x},${point.y}` }
}

function onInkPointerMove(event: PointerEvent) {
  if (activeTool.value !== 'laser' && !event.buttons && event.pointerType !== 'touch') return
  const point = inkPoint(event)
  if (activeTool.value === 'laser') laserPoint.value = point
  else if (draftStroke.value) draftStroke.value = { ...draftStroke.value, points: `${draftStroke.value.points} ${point.x},${point.y}` }
}

function onInkPointerUp() {
  const stroke = draftStroke.value
  if (stroke && currentSlide.value && stroke.points.includes(' ')) {
    inkBySlide.value = { ...inkBySlide.value, [currentSlide.value.id]: [...currentInk.value, stroke] }
  }
  draftStroke.value = null
  laserPoint.value = null
}

function clearInk() {
  if (!currentSlide.value) return
  inkBySlide.value = { ...inkBySlide.value, [currentSlide.value.id]: [] }
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
  const slide = currentSlide.value
  const delayMs = slide?.advanceAfterMs
  if (!slide || delayMs === undefined || currentIndex.value >= (deck.value?.slides.length || 0) - 1) return
  const slideIndex = currentIndex.value
  const dueAt = Date.now() + delayMs
  const advance = () => {
    const remaining = dueAt - Date.now()
    if (remaining > 0) {
      slideAdvanceTimer = setTimeout(advance, Math.min(remaining, 2_147_483_647))
      return
    }
    if (currentIndex.value !== slideIndex || currentSlide.value !== slide) return
    animationStep.value = animationCount.value
    nextSlide()
  }
  slideAdvanceTimer = setTimeout(advance, Math.min(delayMs, 2_147_483_647))
}

watch(currentIndex, (nextIndex, previousIndex) => {
  if (nextIndex === previousIndex) return
  continueMediaAcrossSlides(nextIndex, previousIndex)
  draftStroke.value = null
  laserPoint.value = null
  zoom.value = 1
}, { flush: 'sync' })

watch([() => deck.value, () => currentIndex.value], scheduleSlideAdvance, { immediate: true })

onMounted(() => {
  window.addEventListener('keydown', onKeydown)
  window.addEventListener('resize', scheduleThumbnailUpdate)
  void nextTick(updateThumbnailWindow)
})
onBeforeUnmount(() => {
  fileLoadGeneration++
  window.removeEventListener('keydown', onKeydown)
  window.removeEventListener('resize', scheduleThumbnailUpdate)
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
  <main class="app-shell">
    <header class="topbar">
      <div class="brand" aria-label="Deckline PPTX Player">
        <span class="brand-mark">D</span>
        <span class="brand-name">deckline</span>
        <span class="brand-label">PPTX PLAYER</span>
      </div>
      <div class="topbar-status"><span class="status-dot" />LOCAL SESSION · FILES STAY ON DEVICE</div>
    </header>

    <div class="workspace" :inert="presenterView || sorterView">
      <aside class="deck-sidebar" aria-label="Slide list">
        <div class="sidebar-head">
          <span class="sidebar-title">Slides</span>
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
            :aria-label="`Go to slide ${index + 1}`"
            @click="goToSlide(index)"
          >
            <div class="thumb-preview">
              <SlideStage :slide="slide" :width="deck.width" :height="deck.height" :number="index + 1" thumbnail />
            </div>
            <div class="thumb-meta"><span>SLIDE</span><span>{{ String(index + 1).padStart(2, '0') }}</span></div>
          </button>
          <div v-if="thumbnailEnd < thumbnailCount" class="thumb-virtual-spacer" :style="thumbnailAfterStyle" aria-hidden="true" />
        </div>
        <div v-else class="side-empty">Your slide thumbnails will appear here after you open a presentation.</div>
        <div class="sidebar-footer">POWERED BY OOXML · BROWSER ONLY</div>
      </aside>

      <section class="player-column" aria-label="Presentation player">
        <div class="deck-toolbar">
          <span class="file-icon">▤</span>
          <div class="deck-title-block">
            <div class="deck-title">{{ deck?.name || 'No presentation open' }}</div>
            <div class="deck-subtitle">{{ deck ? `${deck.slides.length} SLIDES · READY TO PRESENT` : 'OPEN A POWERPOINT FILE TO BEGIN' }}</div>
          </div>
          <div class="tool-separator" />
          <button class="icon-button" :disabled="!deck || (currentIndex === 0 && animationStep === 0)" aria-label="Previous animation step or slide" title="Previous animation step or slide" @click="previousSlide">‹</button>
          <button class="icon-button" :disabled="!deck || (currentIndex === (deck?.slides.length || 0) - 1 && animationStep >= animationCount)" aria-label="Next animation step or slide" title="Next animation step or slide" @click="nextSlide">›</button>
          <button class="text-button" :disabled="!deck" title="Fullscreen (F)" @click="toggleFullscreen">⛶ <span>Present</span></button>
          <button ref="sorterTrigger" class="tool-mode-button" data-tool="sorter" :disabled="!deck" aria-label="Open slide sorter" title="Slide sorter (G)" @click="toggleSorterView">sort</button>
          <button ref="presenterTrigger" class="tool-mode-button" data-tool="speaker" :disabled="!deck" aria-label="Open speaker view" title="Speaker view (S)" @click="togglePresenterView">speaker</button>
          <button v-for="tool in (['pointer', 'pen', 'highlighter', 'laser'] as const)" :key="tool" class="tool-mode-button" :class="{ 'is-active': activeTool === tool }" :data-tool="tool" :disabled="!deck" @click="activeTool = tool">{{ tool }}</button>
          <button class="tool-mode-button" :disabled="!deck || !currentInk.length" data-tool="clear" @click="clearInk">clear</button>
          <div class="tool-separator" />
          <button class="primary-button" :disabled="isLoading" @click="openPicker">{{ isLoading ? 'Reading…' : deck ? 'Open another' : 'Open PPTX' }}</button>
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
          <div ref="continuingMediaHost" class="continuing-media-host" aria-hidden="true" />
          <div v-if="isDragging" class="upload-overlay">DROP YOUR PRESENTATION</div>
            <div v-if="currentSlide && deck" ref="stageSlideHost" class="slide-host" :data-slide-number="currentNumber" :data-transition-effect="currentSlide.transition?.effect" :data-transition-preset="currentSlide.transition?.presetName" :style="{ '--slide-ratio': String(deck.width / deck.height), transform: `scale(${zoom})` }">
            <SlideStage :key="deckGeneration + '-' + currentSlide.id" :slide="currentSlide" :width="deck.width" :height="deck.height" :number="currentNumber" :animation-step="animationStep" @trigger-playback="triggerPlayback = $event" @activate-hyperlink="activateHyperlink" />
            <svg
              class="ink-layer"
              :class="{ 'is-active': activeTool !== 'pointer' }"
              :viewBox="`0 0 ${deck.width} ${deck.height}`"
              preserveAspectRatio="none"
              aria-label="Presentation annotation layer"
              @pointerdown.stop.prevent="onInkPointerDown"
              @pointermove.stop.prevent="onInkPointerMove"
              @pointerup.stop.prevent="onInkPointerUp"
              @pointercancel.stop.prevent="onInkPointerUp"
              @pointerleave="activeTool === 'laser' && (laserPoint = null)"
            >
              <polyline v-for="(stroke, index) in currentInk" :key="index" :points="stroke.points" fill="none" :stroke="stroke.tool === 'pen' ? '#ef4444' : '#fde047'" :stroke-width="stroke.tool === 'pen' ? 3 : 16" :opacity="stroke.tool === 'pen' ? 1 : .4" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
              <polyline v-if="draftStroke" :points="draftStroke.points" fill="none" :stroke="draftStroke.tool === 'pen' ? '#ef4444' : '#fde047'" :stroke-width="draftStroke.tool === 'pen' ? 3 : 16" :opacity="draftStroke.tool === 'pen' ? 1 : .4" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
              <circle v-if="laserPoint" :cx="laserPoint.x" :cy="laserPoint.y" :r="deck.width / 120" fill="#ef4444" opacity=".9" />
            </svg>
            <div ref="flashOverlay" class="flash-transition-overlay" aria-hidden="true" />
          </div>
          <div v-else-if="!errorMessage" class="empty-state">
            <div class="empty-art"><span class="empty-art-back" /><span class="empty-art-front" /></div>
            <div class="empty-kicker">A quieter way to present</div>
            <h1>Your deck, in the room.</h1>
            <p>Open a PowerPoint file and step through your slides. The presentation is read locally in your browser.</p>
            <button class="primary-button" :disabled="isLoading" @click.stop="openPicker">{{ isLoading ? 'Reading presentation…' : 'Choose a .pptx file' }}</button>
            <div class="drop-hint">OR DROP A FILE ANYWHERE IN THIS WINDOW</div>
          </div>
          <div v-if="screenOverlay" class="screen-overlay" :class="`is-${screenOverlay}`" @click.stop="screenOverlay = null" />
          <div v-if="slideNumberInput" class="slide-number-input" aria-live="polite">{{ slideNumberInput }}</div>
          <div v-if="contextMenu" class="presentation-menu" :style="{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }" role="menu" @click.stop>
            <button role="menuitem" :disabled="currentIndex === 0 && animationStep === 0" @click="previousSlide(); contextMenu = null">Previous</button>
            <button role="menuitem" :disabled="currentIndex === (deck?.slides.length || 0) - 1 && animationStep >= animationCount" @click="nextSlide(); contextMenu = null">Next</button>
            <button role="menuitem" @click="screenOverlay = 'black'; contextMenu = null">Black screen</button>
            <button role="menuitem" @click="screenOverlay = 'white'; contextMenu = null">White screen</button>
          </div>
          <div v-if="errorMessage" class="error-card" role="alert">
            <strong>Couldn’t open this presentation</strong>{{ errorMessage }}
          </div>
          <div v-if="deck?.warnings.length" class="warning-strip" aria-label="Rendering limitations">
            <span v-for="warning in deck.warnings" :key="warning" class="warning-pill">{{ warning }}</span>
          </div>
          <div v-if="deck" class="progress-track" aria-hidden="true"><div class="progress-value" :style="{ width: `${progress}%` }" /></div>
        </div>

        <div class="player-controls">
          <button class="icon-button" :disabled="!deck || (currentIndex === 0 && animationStep === 0)" aria-label="Previous animation step or slide" @click="previousSlide">←</button>
          <div class="slide-position"><strong>{{ String(currentNumber).padStart(2, '0') }}</strong> / {{ String(deck?.slides.length || 0).padStart(2, '0') }}</div>
          <button class="icon-button" :disabled="!deck || (currentIndex === (deck?.slides.length || 0) - 1 && animationStep >= animationCount)" aria-label="Next animation step or slide" @click="nextSlide">→</button>
          <button class="zoom-button" :disabled="!deck || zoom <= .5" aria-label="Zoom out" @click="setZoom(zoom - .25)">−</button>
          <span class="zoom-value">{{ Math.round(zoom * 100) }}%</span>
          <button class="zoom-button" :disabled="!deck || zoom >= 3" aria-label="Zoom in" @click="setZoom(zoom + .25)">+</button>
          <span class="shortcut-hint">ARROWS · F5 START · SHIFT+F5 CURRENT · F FULLSCREEN</span>
        </div>
      </section>
    </div>
    <div v-if="presenterView && deck && currentSlide" class="presenter-view" data-presenter-view :data-current-slide="currentNumber" :data-animation-step="animationStep" role="dialog" aria-modal="true" aria-label="Speaker view">
      <header class="presenter-header">
        <div><strong>{{ deck.name }}</strong><span> SPEAKER VIEW</span></div>
        <div class="presenter-header-actions">
          <span class="presenter-clock" aria-label="Elapsed time">{{ presenterTime }}</span>
          <span class="presenter-display-status" :data-projector-connected="projectorConnected" :data-projector-number="projectorAcknowledgement?.number ?? 0" :data-projector-step="projectorAcknowledgement?.animationStep ?? -1" :data-projector-blank="projectorAcknowledgement?.blank || ''" :data-projector-trigger-count="projectorAcknowledgement?.triggerCount ?? -1">{{ projectorConnected ? 'SECOND DISPLAY CONNECTED' : 'SECOND DISPLAY CLOSED' }}</span>
          <button class="text-button" data-open-projector @click="openProjectorWindow">{{ projectorOpen ? 'Focus display' : 'Open second display' }}</button>
          <span v-if="projectorError" class="projector-error" data-projector-error role="alert">{{ projectorError }}</span>
          <button v-if="projectorOpen" class="text-button" data-close-projector @click="closeProjectorWindow">Close display</button>
          <button ref="presenterExit" class="text-button" data-presenter-exit @click="togglePresenterView">Exit <kbd>Esc</kbd></button>
        </div>
      </header>
      <div class="presenter-grid">
        <section class="presenter-current-panel" aria-label="Current slide">
          <div class="presenter-panel-title"><span>Current</span><span>{{ currentNumber }} / {{ deck.slides.length }}</span></div>
          <div class="presenter-current-frame">
            <div ref="presenterStageShell" class="slide-transition-shell presenter-transition-shell">
              <div ref="presenterTransitionUnderlay" class="slide-transition-underlay" :class="{ 'is-pulling': currentSlide?.transition?.effect === 'pull', 'is-splitting-in': currentSlide?.transition?.effect === 'split' && currentSlide.transition.direction === 'in' }" aria-hidden="true" />
              <div ref="presenterSlideHost" class="slide-host" :data-slide-number="currentNumber" :style="{ '--slide-ratio': String(deck.width / deck.height) }">
                <SlideStage class="presenter-large-slide" :style="{ '--slide-ratio': String(deck.width / deck.height) }" :slide="currentSlide" :width="deck.width" :height="deck.height" :number="currentNumber" :animation-step="animationStep" :trigger-playback="triggerPlayback" thumbnail playback-preview />
              </div>
              <div ref="presenterFlashOverlay" class="flash-transition-overlay" aria-hidden="true" />
            </div>
          </div>
        </section>
        <aside class="presenter-side-panel">
          <section class="presenter-next-panel" aria-label="Next slide">
            <div class="presenter-panel-title"><span>Next</span><span>{{ currentIndex < deck.slides.length - 1 ? `${currentIndex + 2} / ${deck.slides.length}` : 'End of deck' }}</span></div>
            <div v-if="deck.slides[currentIndex + 1]" class="presenter-next-frame">
              <SlideStage class="presenter-next-slide" :slide="deck.slides[currentIndex + 1]!" :width="deck.width" :height="deck.height" :number="currentIndex + 2" thumbnail />
            </div>
            <div v-else class="presenter-end-card">This is the final slide.</div>
          </section>
          <section class="presenter-notes-panel" aria-label="Speaker notes">
            <div class="presenter-panel-title"><span>Speaker notes</span></div>
            <p v-if="currentSlide.speakerNotes" data-presenter-notes>{{ currentSlide.speakerNotes }}</p>
            <p v-else class="presenter-no-notes" data-presenter-notes>No speaker notes on this slide.</p>
          </section>
        </aside>
      </div>
      <footer class="presenter-footer">
        <button class="text-button" :disabled="currentIndex === 0 && animationStep === 0" @click="previousSlide">← Previous</button>
        <span>Arrow keys navigate · S toggles speaker view · {{ presenterTime }}</span>
        <button class="text-button" :disabled="currentIndex === deck.slides.length - 1 && animationStep >= animationCount" @click="nextSlide">Next →</button>
      </footer>
    </div>
    <div v-if="sorterView && deck" class="sorter-view" data-sorter-view role="dialog" aria-modal="true" aria-label="Slide sorter">
      <header class="sorter-header">
        <div><strong>Slide sorter</strong><span>{{ sorterPageStart + 1 }}–{{ sorterPageEnd }} / {{ deck.slides.length }}</span></div>
        <button ref="sorterExit" class="text-button" data-sorter-exit @click="toggleSorterView">Close <kbd>Esc</kbd></button>
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
          <button class="sorter-slide-button" :aria-label="`Go to slide ${index + 1}: ${slide.name}`" @click="goToSlide(index); toggleSorterView()">
            <div class="sorter-preview" :style="{ aspectRatio: `${deck.width} / ${deck.height}` }"><SlideStage :slide="slide" :width="deck.width" :height="deck.height" :number="index + 1" thumbnail /></div>
            <div class="sorter-meta"><span>{{ String(index + 1).padStart(2, '0') }}</span><span>{{ slide.name }}</span></div>
          </button>
        </article>
      </div>
      <footer class="sorter-footer">
        <button class="text-button" :disabled="sorterPageStart === 0" @click="stepSorterPage(-1)">← Previous group</button>
        <span>Drag to reorder in this group · Select a slide, then use Alt + arrow keys to move it</span>
        <button class="text-button" :disabled="sorterPageEnd >= deck.slides.length" @click="stepSorterPage(1)">Next group →</button>
      </footer>
    </div>
  </main>
</template>
