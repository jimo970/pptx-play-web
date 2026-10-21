import { nextTick, onUnmounted, watch, type Ref } from 'vue'
import type { PptxSlide, PptxTransitionPlayback } from '~/utils/pptx-model'
import { combClipPath, dissolveClipPath, glitterTransitionClipPath, gridTransitionClipPath, plusTransitionClipPath, randomBarClipPath, rippleTransitionClipPath, shredTransitionClipPath, stripsClipPath, wedgeTransitionClipPath, wheelClipPath } from '~/utils/transitionMasks'

interface SlideTransitionOptions {
  slide: Readonly<Ref<PptxSlide | null>>
  shell: Ref<HTMLElement | null>
  incomingHost: Ref<HTMLElement | null>
  underlay: Ref<HTMLElement | null>
  flash: Ref<HTMLElement | null>
  playback: Readonly<Ref<PptxTransitionPlayback | null>>
}

type ConcreteTransitionEffect = Exclude<NonNullable<PptxSlide['transition']>['effect'], 'random' | 'preset'>

const randomTransitionEffects: readonly ConcreteTransitionEffect[] = [
  'fade', 'wipe', 'zoom', 'cover', 'pull', 'split', 'circle', 'diamond', 'plus', 'wedge',
  'newsflash', 'doors', 'window', 'prism', 'pan', 'vortex', 'ferris', 'gallery', 'conveyor', 'ripple',
  'glitter', 'shred', 'flythrough', 'warp', 'flip', 'switch', 'reveal', 'wheel', 'morph', 'cut',
  'honeycomb', 'blinds', 'checker', 'comb', 'dissolve', 'wheelReverse', 'randomBar', 'flash', 'strips', 'push',
]

type SlideTransition = NonNullable<PptxSlide['transition']>

// p:random has no concrete effect attributes, so selected effects need their own schema defaults.
const randomTransitionDefaults: Partial<Record<ConcreteTransitionEffect, Partial<SlideTransition>>> = {
  zoom: { direction: 'in' },
  split: { direction: 'out', orientation: 'horz' },
  doors: { orientation: 'horz' },
  window: { orientation: 'horz' },
  prism: { isInverted: false, isContent: false },
  ripple: { direction: 'center' },
  glitter: { direction: 'l', pattern: 'diamond' },
  shred: { direction: 'in', shredPattern: 'strip' },
  flythrough: { direction: 'in', hasBounce: false },
  warp: { direction: 'out' },
  morph: { morphOption: 'byObject' },
  blinds: { orientation: 'horz' },
  checker: { orientation: 'horz' },
  comb: { orientation: 'horz' },
  wheel: { spokes: 4 },
  wheelReverse: { spokes: 4 },
  randomBar: { barOrientation: 'horizontal' },
  strips: { direction: 'ld' },
}

function resolveRandomTransition(transition: NonNullable<PptxSlide['transition']>) {
  const effect = randomTransitionEffects[Math.floor(Math.random() * randomTransitionEffects.length)] || 'fade'
  return { ...transition, ...randomTransitionDefaults[effect], effect }
}

function honeycombClipPath(width: number, height: number, collapsed: boolean): string {
  const side = Math.sqrt(width * height / 340)
  const halfWidth = Math.sqrt(3) * side / 2
  const rowStep = side * 1.5
  const columnStep = halfWidth * 2
  const vertices: [number, number][] = [[0, -side], [halfWidth, -side / 2], [halfWidth, side / 2], [0, side], [-halfWidth, side / 2], [-halfWidth, -side / 2]]
  const paths: string[] = []
  for (let row = -Math.ceil(height / rowStep) - 2; row <= Math.ceil(height / rowStep) + 2; row++) {
    for (let column = -Math.ceil(width / columnStep) - 2; column <= Math.ceil(width / columnStep) + 2; column++) {
      const x = width / 2 + columnStep * (column + row / 2)
      const y = height / 2 + rowStep * row
      if (x + halfWidth < 0 || x - halfWidth > width || y + side < 0 || y - side > height) continue
      paths.push(`M${x.toFixed(2)} ${y.toFixed(2)} ${vertices.map(([dx, dy]) => `L${(x + (collapsed ? 0 : dx)).toFixed(2)} ${(y + (collapsed ? 0 : dy)).toFixed(2)}`).join(' ')} Z`)
    }
  }
  return `path("${paths.join(' ')}")`
}

interface MorphTextToken {
  text: string
  key: string
  left: number
  top: number
  width: number
  height: number
  content: DocumentFragment
  style: Pick<CSSStyleDeclaration, 'color' | 'fontFamily' | 'fontSize' | 'fontStyle' | 'fontWeight' | 'fontVariant' | 'fontStretch' | 'lineHeight' | 'letterSpacing' | 'wordSpacing' | 'textDecoration' | 'textShadow' | 'textTransform' | 'direction' | 'writingMode' | 'verticalAlign'>
}

function withoutMorphTransforms<T>(elements: HTMLElement[], measure: () => T): T {
  const styles = new Map<HTMLElement, [string, string]>()
  for (const element of elements) {
    for (let current: HTMLElement | null = element; current; current = current.parentElement) {
      if (getComputedStyle(current).transform !== 'none' && !styles.has(current)) {
        styles.set(current, [current.style.getPropertyValue('transform'), current.style.getPropertyPriority('transform')])
        current.style.setProperty('transform', 'none', 'important')
      }
      if (current.classList.contains('slide-frame')) break
    }
  }
  try {
    return measure()
  } finally {
    for (const [element, [value, priority]] of styles) {
      if (value) element.style.setProperty('transform', value, priority)
      else element.style.removeProperty('transform')
    }
  }
}

function morphTextSegments(text: string, mode: 'byWord' | 'byChar'): {text: string; key: string; start: number; end: number}[] {
  if (typeof Intl.Segmenter !== 'function') {
    if (mode === 'byChar') {
      let offset = 0
      return Array.from(text, segment => {
        const start = offset
        offset += segment.length
        return {text: segment, key: segment, start, end: offset}
      })
    }
    return [...text.matchAll(/\S+\s*|\s+/gu)].map(match => ({text: match[0], key: match[0].trimEnd() || match[0], start: match.index!, end: match.index! + match[0].length}))
  }

  const segments = [...new Intl.Segmenter(undefined, {granularity: mode === 'byChar' ? 'grapheme' : 'word'}).segment(text)]
  if (mode === 'byChar') return segments.map(segment => ({text: segment.segment, key: segment.segment, start: segment.index, end: segment.index + segment.segment.length}))
  const words = segments.filter(segment => segment.isWordLike)
  if (!words.length) return text ? [{text, key: text, start: 0, end: text.length}] : []
  return words.map((word, index) => {
    const start = index === 0 ? 0 : word.index
    const end = index + 1 < words.length ? words[index + 1]!.index : text.length
    return {text: text.slice(start, end), key: word.segment, start, end}
  })
}

function morphTextTokens(frame: HTMLElement, mode: 'byWord' | 'byChar'): MorphTextToken[] {
  const tokens: MorphTextToken[] = []
  const frameRect = frame.getBoundingClientRect()
  for (const paragraph of frame.querySelectorAll<HTMLElement>('.text-paragraph')) {
    const nodes: {node: Text; start: number; end: number}[] = []
    let text = ''
    for (const run of paragraph.querySelectorAll<HTMLElement>('.text-run')) {
      const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT)
      let node = walker.nextNode() as Text | null
      while (node) {
        const start = text.length
        text += node.data
        nodes.push({node, start, end: text.length})
        node = walker.nextNode() as Text | null
      }
    }
    for (const segment of morphTextSegments(text, mode)) {
      const start = nodes.find(item => segment.start >= item.start && segment.start < item.end)
      const end = [...nodes].reverse().find(item => segment.end > item.start && segment.end <= item.end)
      if (!start || !end) continue
      const range = document.createRange()
      range.setStart(start.node, segment.start - start.start)
      range.setEnd(end.node, segment.end - end.start)
      const rect = range.getBoundingClientRect()
      if (!rect.width || !rect.height) continue
      const computed = getComputedStyle(start.node.parentElement || frame)
      tokens.push({
        text: segment.text,
        key: segment.key,
        left: rect.left - frameRect.left,
        top: rect.top - frameRect.top,
        width: rect.width,
        height: rect.height,
        content: range.cloneContents(),
        style: {
          color: computed.color,
          fontFamily: computed.fontFamily,
          fontSize: computed.fontSize,
          fontStyle: computed.fontStyle,
          fontWeight: computed.fontWeight,
          fontVariant: computed.fontVariant,
          fontStretch: computed.fontStretch,
          lineHeight: computed.lineHeight,
          letterSpacing: computed.letterSpacing,
          wordSpacing: computed.wordSpacing,
          textDecoration: computed.textDecoration,
          textShadow: computed.textShadow,
          textTransform: computed.textTransform,
          direction: computed.direction,
          writingMode: computed.writingMode,
          verticalAlign: computed.verticalAlign,
        },
      })
    }
  }
  return tokens
}

function morphTextNode(layer: HTMLElement, token: MorphTextToken, left: number, top: number): HTMLElement {
  const node = document.createElement('span')
  node.className = 'morph-text-token'
  node.dataset.morphToken = token.text
  node.dataset.morphMatch = token.key
  node.setAttribute('aria-hidden', 'true')
  Object.assign(node.style, {
    position: 'absolute',
    display: 'block',
    left: `${left}px`,
    top: `${top}px`,
    width: `${Math.max(token.width, 0.01)}px`,
    height: `${Math.max(token.height, 0.01)}px`,
    margin: '0',
    padding: '0',
    border: '0',
    overflow: 'visible',
    whiteSpace: 'pre',
    pointerEvents: 'none',
    transformOrigin: 'top left',
    ...token.style,
  })
  node.append(token.content)
  for (const run of node.querySelectorAll<HTMLElement>('.text-run')) run.style.setProperty('white-space', 'pre', 'important')
  layer.append(node)
  return node
}


export function useSlideTransition({ slide, shell, incomingHost, underlay, flash, playback }: SlideTransitionOptions) {
  let morphCleanup: (() => void) | undefined

  onUnmounted(() => {
    morphCleanup?.()
    morphCleanup = undefined
  })

  watch(slide, (nextSlide, previousSlide) => {
    if (nextSlide?.id === previousSlide?.id) return
    morphCleanup?.()
    morphCleanup = undefined
    underlay.value?.replaceChildren()
    if (!nextSlide) return
    const current = incomingHost.value
    if (!current) return
    const snapshot = current.cloneNode(true) as HTMLElement
    const style = getComputedStyle(current)
    snapshot.setAttribute('aria-hidden', 'true')
    snapshot.inert = true
    snapshot.style.pointerEvents = 'none'
    snapshot.style.transform = style.transform
    snapshot.style.opacity = style.opacity
    snapshot.style.clipPath = style.clipPath
    const currentContent = current.querySelector<HTMLElement>('.slide-content-layer')
    const snapshotContent = snapshot.querySelector<HTMLElement>('.slide-content-layer')
    if (currentContent && snapshotContent) {
      const contentStyle = getComputedStyle(currentContent)
      snapshotContent.style.transform = contentStyle.transform
      snapshotContent.style.opacity = contentStyle.opacity
    }
    const currentBackground = current.querySelector<HTMLElement>('.slide-background-layer')
    const snapshotBackground = snapshot.querySelector<HTMLElement>('.slide-background-layer')
    if (currentBackground && snapshotBackground) snapshotBackground.style.opacity = getComputedStyle(currentBackground).opacity
    underlay.value?.append(snapshot)
  }, { flush: 'sync' })

  watch(slide, async (activeSlide, previousSlide) => {
    if (!activeSlide || activeSlide.id === previousSlide?.id) return
    const currentPlayback = playback.value?.slideId === activeSlide.id ? playback.value : null
    const sourceTransition = currentPlayback?.transition
    const transition = sourceTransition?.effect === 'random' ? resolveRandomTransition(sourceTransition) : sourceTransition
    await nextTick()
    if (slide.value?.id !== activeSlide.id) return
    const animationDelay = currentPlayback ? Math.min(0, currentPlayback.startedAtEpoch - (performance.timeOrigin + performance.now())) : 0
    shell.value?.classList.remove('is-fading-through-black', 'is-cut-through-black', 'is-prism-transition', 'is-flip-transition', 'is-switch-transition', 'is-ferris-transition', 'is-gallery-transition', 'is-conveyor-transition', 'is-morph-transition', 'is-preset-transition')
    underlay.value?.classList.remove('is-shredding-out', 'is-flythrough', 'is-warping-out', 'is-vortexing')
    const incoming = incomingHost.value
    incoming?.getAnimations().forEach(animation => animation.cancel())
    flash.value?.getAnimations().forEach(animation => animation.cancel())
    if (!transition || window.matchMedia('(prefers-reduced-motion: reduce)').matches || !incoming || !underlay.value?.firstElementChild) {
      underlay.value?.replaceChildren()
      return
    }
  if (transition.effect === 'morph') {
    const outgoing = underlay.value?.firstElementChild as HTMLElement | null
    const stage = shell.value
    if (!outgoing || !stage) {
      underlay.value?.replaceChildren()
      return
    }
    const timing: KeyframeAnimationOptions = { duration: transition.durationMs, easing: 'ease-in-out', fill: 'both', delay: animationDelay }
    const animations: Animation[] = []
    const textMode = transition.morphOption === 'byWord' || transition.morphOption === 'byChar' ? transition.morphOption : undefined
    stage.classList.add('is-morph-transition')
    const outgoingFrame = outgoing.querySelector<HTMLElement>('.slide-frame')
    const incomingFrame = incoming.querySelector<HTMLElement>('.slide-frame')
    const outgoingElements = [...(outgoingFrame?.querySelectorAll<HTMLElement>('.slide-element') || [])]
    const incomingElements = [...(incomingFrame?.querySelectorAll<HTMLElement>('.slide-element') || [])]
    const uniqueByName = (elements: HTMLElement[]) => {
      const grouped = new Map<string, HTMLElement[]>()
      for (const element of elements) {
        const name = element.dataset.elementName
        const kind = element.dataset.elementKind
        if (!name || !kind) continue
        const key = `${kind}:${name}`
        grouped.set(key, [...(grouped.get(key) || []), element])
      }
      return new Map([...grouped].filter(([, matches]) => matches.length === 1).map(([key, matches]) => [key, matches[0]!] as const))
    }
    const oldByName = uniqueByName(outgoingElements)
    const newByName = uniqueByName(incomingElements)
    const restoreTextStyles: (() => void)[] = []
    if (textMode && incomingFrame) {
      for (const [key, element] of newByName) {
        const previous = oldByName.get(key)
        if (!previous) continue
        const sourceText = previous.querySelector<HTMLElement>('.text-frame')
        const targetText = element.querySelector<HTMLElement>('.text-frame')
        if (!sourceText || !targetText) continue
        if (element.dataset.elementKind !== 'shape') {
          console.warn(`Morph ${textMode} text matching is not supported on ${element.dataset.elementKind} elements; object motion is used.`)
          continue
        }
        const [sourceTokens, targetTokens] = withoutMorphTransforms([previous, element], () => [
          morphTextTokens(sourceText, textMode),
          morphTextTokens(targetText, textMode),
        ])
        if (!sourceTokens.length && !targetTokens.length) continue
        const position = targetText.style.getPropertyValue('position')
        const positionPriority = targetText.style.getPropertyPriority('position')
        targetText.style.setProperty('position', 'relative')
        restoreTextStyles.push(() => position ? targetText.style.setProperty('position', position, positionPriority) : targetText.style.removeProperty('position'))
        const layer = document.createElement('span')
        layer.className = 'morph-text-layer'
        layer.setAttribute('aria-hidden', 'true')
        Object.assign(layer.style, {position: 'absolute', display: 'block', inset: '0', overflow: 'hidden', pointerEvents: 'none', zIndex: '2'})
        targetText.append(layer)

        const available = new Map<string, {tokens: MorphTextToken[]; index: number}>()
        for (const token of targetTokens) {
          const bucket = available.get(token.key) || {tokens: [], index: 0}
          bucket.tokens.push(token)
          available.set(token.key, bucket)
        }
        const matchedTargets = new Set<MorphTextToken>()
        for (const token of sourceTokens) {
          const matches = available.get(token.key)
          if (matches && matches.index < matches.tokens.length) {
            const target = matches.tokens[matches.index++]!
            matchedTargets.add(target)
            const node = morphTextNode(layer, target, token.left, token.top)
            const scaleX = token.width / Math.max(target.width, 0.01)
            const scaleY = token.height / Math.max(target.height, 0.01)
            animations.push(node.animate([
              {left: `${token.left}px`, top: `${token.top}px`, transform: `scale(${scaleX}, ${scaleY})`},
              {left: `${target.left}px`, top: `${target.top}px`, transform: 'scale(1, 1)'},
            ], timing))
          } else {
            const node = morphTextNode(layer, token, token.left, token.top)
            animations.push(node.animate([{opacity: 1}, {opacity: 0}], timing))
          }
        }
        for (const token of targetTokens) {
          if (matchedTargets.has(token)) continue
          const node = morphTextNode(layer, token, token.left, token.top)
          animations.push(node.animate([{opacity: 0}, {opacity: 1}], timing))
        }
        for (const run of sourceText.querySelectorAll<HTMLElement>('.text-run')) run.style.setProperty('visibility', 'hidden', 'important')
        for (const run of targetText.querySelectorAll<HTMLElement>('.text-run')) {
          const visibility = run.style.getPropertyValue('visibility')
          const priority = run.style.getPropertyPriority('visibility')
          run.style.setProperty('visibility', 'hidden', 'important')
          restoreTextStyles.push(() => visibility ? run.style.setProperty('visibility', visibility, priority) : run.style.removeProperty('visibility'))
        }
      }
    }
    for (const element of incomingElements) {
      const key = `${element.dataset.elementKind}:${element.dataset.elementName}`
      const previous = oldByName.get(key)
      const finalOpacity = getComputedStyle(element).opacity
      if (!previous || !newByName.has(key)) {
        animations.push(element.animate([{ opacity: 0 }, { opacity: finalOpacity }], timing))
        continue
      }
      const color = (node: HTMLElement, property: 'backgroundColor' | 'borderColor' | 'borderWidth') => getComputedStyle(node)[property]
      animations.push(element.animate([
        {
          left: previous.style.left,
          top: previous.style.top,
          width: previous.style.width,
          height: previous.style.height,
          transform: previous.style.transform || 'none',
          transformOrigin: previous.style.transformOrigin || '50% 50%',
          opacity: getComputedStyle(previous).opacity,
          backgroundColor: color(previous, 'backgroundColor'),
          borderColor: color(previous, 'borderColor'),
          borderWidth: color(previous, 'borderWidth'),
        },
        {
          left: element.style.left,
          top: element.style.top,
          width: element.style.width,
          height: element.style.height,
          transform: element.style.transform || 'none',
          transformOrigin: element.style.transformOrigin || '50% 50%',
          opacity: finalOpacity,
          backgroundColor: color(element, 'backgroundColor'),
          borderColor: color(element, 'borderColor'),
          borderWidth: color(element, 'borderWidth'),
        },
      ], timing))
    }
    for (const element of outgoingElements) {
      animations.push(element.animate([{ opacity: getComputedStyle(element).opacity }, { opacity: 0 }], timing))
    }
    if (incomingFrame) {
      const backgroundAnimation = incomingFrame.animate([
        { backgroundColor: 'rgba(0, 0, 0, 0)' },
        { backgroundColor: getComputedStyle(incomingFrame).backgroundColor },
      ], timing)
      animations.push(backgroundAnimation)
      let cleaned = false
      const cleanup = () => {
        if (cleaned) return
        cleaned = true
        animations.forEach(animation => animation.cancel())
        restoreTextStyles.forEach(restore => restore())
        incomingFrame.querySelectorAll('.morph-text-layer').forEach(layer => layer.remove())
        if (morphCleanup === cleanup) morphCleanup = undefined
      }
      morphCleanup = cleanup
      backgroundAnimation.onfinish = () => {
        if (slide.value?.id === activeSlide.id) {
          cleanup()
          underlay.value?.replaceChildren()
          stage.classList.remove('is-morph-transition')
        }
      }
    }
    return
  }
  if (transition.effect === 'preset') {
    const outgoing = underlay.value?.firstElementChild as HTMLElement | null
    const stage = shell.value
    if (!outgoing || !stage || !transition.presetName) {
      underlay.value?.replaceChildren()
      return
    }
    const x = transition.invertX ? -1 : 1
    const y = transition.invertY ? -1 : 1
    const origin = x > 0 ? 'left center' : 'right center'
    const frames: Keyframe[] = (() => {
      switch (transition.presetName) {
        case 'fallOver': return [{ transform: 'perspective(1200px) rotateX(0) rotateZ(0)', opacity: 1 }, { transform: `perspective(1200px) rotateX(${72 * y}deg) rotateZ(${5 * x}deg) translateY(${10 * y}%)`, opacity: 0 }]
        case 'drape': return [{ transform: 'perspective(1200px) rotateX(0) scaleY(1)', opacity: 1 }, { transform: `perspective(1200px) rotateX(${48 * y}deg) scaleY(.9)`, opacity: .7, offset: .55 }, { transform: `perspective(1200px) rotateX(${78 * y}deg) scaleY(.7)`, opacity: 0 }]
        case 'curtains': return [{ transform: 'scaleX(1) rotateY(0)', transformOrigin: origin, opacity: 1 }, { transform: `scaleX(.02) rotateY(${8 * y}deg)`, transformOrigin: origin, opacity: 0 }]
        case 'wind': return [{ transform: 'translate(0, 0) skewX(0)', opacity: 1 }, { transform: `translate(${125 * x}%, ${-110 * y}%) skewX(${14 * y}deg)`, opacity: 0 }]
        case 'prestige': return [{ transform: 'scale(1) rotate(0)', opacity: 1 }, { transform: `scale(.62) rotate(${5 * x}deg) translate(${5 * x}%, ${12 * y}%)`, opacity: 0 }]
        case 'fracture': return [{ clipPath: 'inset(0)', transform: 'scale(1) rotate(0)', opacity: 1 }, { clipPath: `polygon(50% ${50 - 4 * y}%, ${50 + 4 * x}% 50%, 50% ${50 + 4 * y}%, ${50 - 4 * x}% 50%)`, transform: `scale(.82) rotate(${8 * x}deg)`, opacity: 0 }]
        case 'crush': return [{ transform: 'scale(1)', opacity: 1 }, { transform: 'scaleX(.04)', opacity: 0 }]
        case 'peelOff': return [{ transform: 'perspective(1200px) rotateY(0)', transformOrigin: origin, opacity: 1 }, { transform: `perspective(1200px) rotateY(${105 * x}deg)`, transformOrigin: origin, opacity: 0 }]
        case 'pageCurlDouble': return [{ transform: 'perspective(1200px) rotateY(0) scaleX(1)', transformOrigin: 'center center', opacity: 1 }, { transform: `perspective(1200px) rotateY(${82 * x}deg) scaleX(.28)`, transformOrigin: 'center center', opacity: .8, offset: .55 }, { transform: `perspective(1200px) rotateY(${165 * x}deg) scaleX(.04)`, transformOrigin: 'center center', opacity: 0 }]
        case 'pageCurlSingle': return [{ transform: 'perspective(1200px) rotateY(0)', transformOrigin: origin, opacity: 1 }, { transform: `perspective(1200px) rotateY(${112 * x}deg)`, transformOrigin: origin, opacity: 0 }]
        case 'airplane': return [{ transform: 'perspective(1200px) translate(0, 0) rotate(0) rotateY(0) scale(1)', opacity: 1 }, { transform: `perspective(1200px) translate(${130 * x}%, ${95 * y}%) rotate(${18 * x}deg) rotateY(${42 * y}deg) scale(.72)`, opacity: 0 }]
        case 'origami': return [{ transform: 'perspective(1200px) rotateX(0) rotateY(0) scale(1)', opacity: 1 }, { transform: `perspective(1200px) rotateX(${24 * y}deg) rotateY(${34 * x}deg) scale(.88)`, opacity: .9, offset: .32 }, { transform: `perspective(1200px) rotateX(${-32 * y}deg) rotateY(${-42 * x}deg) scale(.68)`, opacity: .65, offset: .68 }, { transform: `perspective(1200px) rotateX(${90 * y}deg) rotateY(${100 * x}deg) translate(${30 * x}%, ${30 * y}%) scale(.48)`, opacity: 0 }]
      }
    })()
    stage.classList.add('is-preset-transition')
    const animation = outgoing.animate(frames, { duration: transition.durationMs, easing: 'ease-in-out', fill: 'both', delay: animationDelay })
    animation.onfinish = () => {
      animation.cancel()
      if (slide.value?.id === activeSlide.id) {
        underlay.value?.replaceChildren()
        stage.classList.remove('is-preset-transition')
      }
    }
    return
  }
  if (transition.effect === 'cut') {
    underlay.value?.replaceChildren()
    if (transition.throughBlack) {
      const stage = shell.value
      stage?.classList.add('is-cut-through-black')
      requestAnimationFrame(() => requestAnimationFrame(() => {
        if (slide.value?.id === activeSlide.id) stage?.classList.remove('is-cut-through-black')
      }))
    }
    return
  }
  const direction = transition.direction
  const effect = transition.effect
  if (['prism', 'flip', 'switch', 'ferris', 'gallery', 'conveyor'].includes(effect)) {
    const outgoing = underlay.value?.firstElementChild as HTMLElement | null
    const stage = shell.value
    if (!outgoing || !stage) {
      underlay.value?.replaceChildren()
      return
    }
    const separatePrismContent = effect === 'prism' && transition.isContent === true
    const incomingPlane = separatePrismContent ? incoming.querySelector<HTMLElement>('.slide-content-layer') : incoming
    const outgoingPlane = separatePrismContent ? outgoing.querySelector<HTMLElement>('.slide-content-layer') : outgoing
    const incomingBackground = separatePrismContent ? incoming.querySelector<HTMLElement>('.slide-background-layer') : null
    const outgoingBackground = separatePrismContent ? outgoing.querySelector<HTMLElement>('.slide-background-layer') : null
    if (!incomingPlane || !outgoingPlane || (separatePrismContent && (!incomingBackground || !outgoingBackground))) {
      underlay.value?.replaceChildren()
      return
    }
    const rotations: Record<string, { axis: string; angle: number }> = {
      l: { axis: '0, 1, 0', angle: 90 }, r: { axis: '0, 1, 0', angle: -90 },
      u: { axis: '1, 0, 0', angle: -90 }, d: { axis: '1, 0, 0', angle: 90 },
      ld: { axis: '1, 1, 0', angle: 90 }, lu: { axis: '1, -1, 0', angle: -90 },
      rd: { axis: '1, -1, 0', angle: 90 }, ru: { axis: '1, 1, 0', angle: -90 },
    }
    const rotation = effect === 'prism'
      ? rotations[direction] || rotations.l!
      : { axis: '0, 1, 0', angle: (direction === 'r' ? 1 : -1) * (effect === 'flip' ? 180 : 90) }
    const angle = rotation.angle * (effect === 'prism' && transition.isInverted ? -1 : 1)
    const incomingBase = incomingPlane.style.transform || 'none'
    const outgoingBase = outgoingPlane.style.transform || 'none'
    const base = (transform: string) => transform === 'none' ? '' : `${transform} `
    const side = direction === 'r' ? -1 : 1
    const startTransform = effect === 'ferris'
      ? `${base(incomingBase)}perspective(1200px) rotateZ(${side * 180}deg) scale(0.78)`
      : effect === 'gallery'
        ? `${base(incomingBase)}translateX(${side * 36}%) perspective(1200px) rotateY(${side * 38}deg) scale(0.84)`
        : effect === 'conveyor'
          ? `${base(incomingBase)}translateX(${side * 100}%) perspective(1200px) rotateX(-10deg) rotateY(${side * 12}deg)`
          : `${base(incomingBase)}rotate3d(${rotation.axis}, ${angle}deg)`
    const endTransform = effect === 'ferris'
      ? `${base(outgoingBase)}perspective(1200px) rotateZ(${side * -180}deg) scale(0.78)`
      : effect === 'gallery'
        ? `${base(outgoingBase)}translateX(${side * -36}%) perspective(1200px) rotateY(${side * -38}deg) scale(0.84)`
        : effect === 'conveyor'
          ? `${base(outgoingBase)}translateX(${side * -100}%) perspective(1200px) rotateX(10deg) rotateY(${side * -12}deg)`
          : `${base(outgoingBase)}rotate3d(${rotation.axis}, ${-angle}deg)`
    const transitionClass = effect === 'prism' ? 'is-prism-transition' : effect === 'flip' ? 'is-flip-transition' : effect === 'switch' ? 'is-switch-transition' : `is-${effect}-transition`
    stage.classList.add(transitionClass)
    const timing: KeyframeAnimationOptions = { duration: transition.durationMs, easing: 'ease-in-out', fill: 'both', delay: animationDelay }
    const incomingAnimation = incomingPlane.animate([{ transform: startTransform }, { transform: incomingBase }], timing)
    const outgoingAnimation = outgoingPlane.animate([{ transform: outgoingBase }, { transform: endTransform }], timing)
    const backgroundAnimations = separatePrismContent
      ? [incomingBackground!.animate([{ opacity: 0 }, { opacity: 1 }], timing), outgoingBackground!.animate([{ opacity: getComputedStyle(outgoingBackground!).opacity }, { opacity: 0 }], timing)]
      : []
    incomingAnimation.onfinish = () => {
      incomingAnimation.cancel()
      outgoingAnimation.cancel()
      backgroundAnimations.forEach(animation => animation.cancel())
      if (slide.value?.id === activeSlide.id) {
        underlay.value?.replaceChildren()
        stage.classList.remove(transitionClass)
      }
    }
    return
  }
  if (effect === 'flash') {
    const animation = flash.value?.animate(
      [{ opacity: 0 }, { opacity: 1, offset: 0.12 }, { opacity: 1, offset: 0.28 }, { opacity: 0 }],
      { duration: transition.durationMs, easing: 'linear', fill: 'both', delay: animationDelay },
    )
    if (animation) animation.onfinish = () => animation.cancel()
    underlay.value?.replaceChildren()
    return
  }
  const revealThroughBlack = effect === 'reveal' && transition.throughBlack
  let frames: Keyframe[]
  let outgoingTransform = ''
  let outgoingFrames: Keyframe[] | undefined
  let shredOutgoing = false
  if (effect === 'fade') {
    frames = transition.throughBlack
      ? [{ opacity: 0, offset: 0 }, { opacity: 0, offset: 0.5 }, { opacity: 1, offset: 1 }]
      : [{ opacity: 0 }, { opacity: 1 }]
  } else if (effect === 'cover') {
    const starts: Record<string, string> = {
      l: 'translateX(100%)', r: 'translateX(-100%)', u: 'translateY(100%)', d: 'translateY(-100%)',
      ld: 'translate(100%, -100%)', lu: 'translate(100%, 100%)', rd: 'translate(-100%, -100%)', ru: 'translate(-100%, 100%)',
    }
    frames = [{ transform: starts[direction] || starts.l }, { transform: 'translate(0, 0)' }]
  } else if (effect === 'push' || effect === 'reveal' || effect === 'pan') {
    const panStarts: Record<string, string> = {
      l: 'translateX(100%)', r: 'translateX(-100%)', u: 'translateY(100%)', d: 'translateY(-100%)',
      ld: 'translate(100%, -100%)', lu: 'translate(100%, 100%)', rd: 'translate(-100%, -100%)', ru: 'translate(-100%, 100%)',
    }
    const panExits: Record<string, string> = {
      l: 'translateX(-100%)', r: 'translateX(100%)', u: 'translateY(-100%)', d: 'translateY(100%)',
      ld: 'translate(-100%, 100%)', lu: 'translate(-100%, -100%)', rd: 'translate(100%, 100%)', ru: 'translate(100%, -100%)',
    }
    const start = effect === 'pan' ? panStarts[direction] || panStarts.l! : direction === 'u' ? 'translateY(100%)' : direction === 'd' ? 'translateY(-100%)' : direction === 'r' ? 'translateX(-100%)' : 'translateX(100%)'
    outgoingTransform = effect === 'pan' ? panExits[direction] || panExits.l! : direction === 'u' ? 'translateY(-100%)' : direction === 'd' ? 'translateY(100%)' : direction === 'r' ? 'translateX(100%)' : 'translateX(-100%)'
    frames = revealThroughBlack
      ? [{ transform: start, opacity: 0 }, { transform: start, opacity: 0, offset: 0.5 }, { transform: 'translate(0, 0)', opacity: 1 }]
      : [{ transform: start }, { transform: 'translate(0, 0)' }]
  } else if (effect === 'pull') {
    frames = []
  } else if (effect === 'split') {
    const center = transition.orientation === 'horz' ? 'inset(50% 0)' : 'inset(0 50%)'
    frames = transition.direction === 'in' ? [] : [{ clipPath: center }, { clipPath: 'inset(0)' }]
  } else if (effect === 'doors' || effect === 'window') {
    const center = effect === 'window'
      ? 'inset(50%)'
      : transition.orientation === 'vert' ? 'inset(0 50%)' : 'inset(50% 0)'
    frames = [{ clipPath: center }, { clipPath: 'inset(0)' }]
  } else if (effect === 'honeycomb') {
    const { width, height } = incoming.getBoundingClientRect()
    frames = [{ clipPath: honeycombClipPath(width, height, true) }, { clipPath: honeycombClipPath(width, height, false) }]
  } else if (effect === 'strips') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: stripsClipPath(width, height, direction, progress),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'dissolve') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: dissolveClipPath(width, height, progress, activeSlide.id, 16),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'comb') {
    const { width, height } = incoming.getBoundingClientRect()
    const orientation = transition.orientation || 'horz'
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: combClipPath(width, height, orientation, progress),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'blinds' || effect === 'checker') {
    const { width, height } = incoming.getBoundingClientRect()
    const orientation = transition.orientation || 'horz'
    const progressValues = effect === 'checker' ? [0, 0.25, 0.5, 0.75, 1] : [0, 0.5, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: gridTransitionClipPath(width, height, orientation, progress, effect === 'checker'),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'randomBar') {
    const { width, height } = incoming.getBoundingClientRect()
    const orientation = transition.barOrientation || 'horizontal'
    const seed = activeSlide.id
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: randomBarClipPath(width, height, orientation, progress, seed),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'wheel' || effect === 'wheelReverse') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: wheelClipPath(width, height, transition.spokes || 4, progress, effect === 'wheelReverse'),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'plus' || effect === 'wedge') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: effect === 'plus'
        ? plusTransitionClipPath(width, height, progress)
        : wedgeTransitionClipPath(width, height, progress),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'ripple') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.2, 0.4, 0.6, 0.8, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: rippleTransitionClipPath(width, height, transition.direction as 'center' | 'lu' | 'ru' | 'ld' | 'rd', progress),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'glitter') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    frames = progressValues.map((progress, index) => ({
      clipPath: glitterTransitionClipPath(width, height, direction, transition.pattern || 'diamond', progress),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'shred') {
    const { width, height } = incoming.getBoundingClientRect()
    const progressValues = [0, 0.25, 0.5, 0.75, 1]
    shredOutgoing = transition.direction === 'out'
    if (shredOutgoing) underlay.value?.classList.add('is-shredding-out')
    frames = progressValues.map((progress, index) => ({
      clipPath: shredTransitionClipPath(width, height, transition.shredPattern || 'strip', shredOutgoing ? 1 - progress : progress),
      offset: index / (progressValues.length - 1),
    }))
  } else if (effect === 'flythrough') {
    const baseTransform = incoming.style.transform === 'none' ? '' : `${incoming.style.transform} `
    const outgoing = underlay.value?.firstElementChild as HTMLElement | null
    const outgoingBase = outgoing?.style.transform && outgoing.style.transform !== 'none' ? `${outgoing.style.transform} ` : ''
    const incomingTransform = (scale: number) => `${baseTransform}scale(${scale})`.trim()
    const outgoingTransformAt = (scale: number) => `${outgoingBase}scale(${scale})`.trim()
    underlay.value?.classList.add('is-flythrough')
    frames = transition.hasBounce
      ? [{ transform: incomingTransform(direction === 'in' ? 2.4 : 0.35), opacity: 0 }, { transform: incomingTransform(1.08), opacity: 1, offset: 0.82 }, { transform: incomingTransform(1), opacity: 1 }]
      : [{ transform: incomingTransform(direction === 'in' ? 2.4 : 0.35), opacity: 0 }, { transform: incomingTransform(1), opacity: 1 }]
    outgoingFrames = direction === 'in'
      ? [{ transform: outgoingTransformAt(1), opacity: 1 }, { transform: outgoingTransformAt(0.35), opacity: 0 }]
      : [{ transform: outgoingTransformAt(1), opacity: 1 }, { transform: outgoingTransformAt(2.4), opacity: 0 }]
  } else if (effect === 'warp') {
    const incomingBase = incoming.style.transform && incoming.style.transform !== 'none' ? `${incoming.style.transform} ` : ''
    const outgoing = underlay.value?.firstElementChild as HTMLElement | null
    const outgoingBase = outgoing?.style.transform && outgoing.style.transform !== 'none' ? `${outgoing.style.transform} ` : ''
    const warped = (base: string, x: number, y: number, scale: number) => `${base}perspective(1200px) rotateX(${x}deg) rotateY(${y}deg) skewX(${y / 3}deg) scale(${scale})`.trim()
    if (direction === 'out') {
      underlay.value?.classList.add('is-warping-out')
      frames = [{ transform: incomingBase.trim() || 'none', opacity: 1 }, { transform: incomingBase.trim() || 'none', opacity: 1 }]
      outgoingFrames = [
        { transform: outgoingBase.trim() || 'none', opacity: 1 },
        { transform: warped(outgoingBase, 18, -24, 0.72), opacity: 0, offset: 1 },
      ]
    } else {
      frames = [
        { transform: warped(incomingBase, -18, 24, 0.72), opacity: 0 },
        { transform: warped(incomingBase, 5, -7, 1.04), opacity: 1, offset: 0.78 },
        { transform: incomingBase.trim() || 'none', opacity: 1 },
      ]
    }
  } else if (effect === 'vortex') {
    const outgoing = underlay.value?.firstElementChild as HTMLElement | null
    const outgoingBase = outgoing?.style.transform && outgoing.style.transform !== 'none' ? `${outgoing.style.transform} ` : ''
    const rotations: Record<string, { axis: string; angle: number }> = {
      l: { axis: '0, 0, 1', angle: -540 }, r: { axis: '0, 0, 1', angle: 540 },
      u: { axis: '1, 0, 0', angle: -540 }, d: { axis: '1, 0, 0', angle: 540 },
      ld: { axis: '1, 1, 0', angle: -540 }, lu: { axis: '1, -1, 0', angle: 540 },
      rd: { axis: '1, -1, 0', angle: 540 }, ru: { axis: '1, 1, 0', angle: -540 },
    }
    const rotation = rotations[direction] || rotations.l!
    const endTransform = `${outgoingBase}perspective(1000px) rotate3d(${rotation.axis}, ${rotation.angle}deg) scale(0.08)`.trim()
    underlay.value?.classList.add('is-vortexing')
    frames = [{ transform: incoming.style.transform || 'none', opacity: 1 }, { transform: incoming.style.transform || 'none', opacity: 1 }]
    outgoingFrames = [{ transform: outgoingBase.trim() || 'none', opacity: 1 }, { transform: endTransform, opacity: 0 }]
  } else if (effect === 'circle') {
    frames = [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(100% at 50% 50%)' }]
  } else if (effect === 'diamond') {
    frames = [
      { clipPath: 'polygon(50% 50%, 50% 50%, 50% 50%, 50% 50%)' },
      { clipPath: 'polygon(50% -50%, 150% 50%, 50% 150%, -50% 50%)' },
    ]
  } else if (effect === 'newsflash') {
    frames = [{ transform: 'scale(.2) rotate(360deg)', opacity: 0 }, { transform: 'scale(1) rotate(0deg)', opacity: 1 }]
  } else if (effect === 'zoom') {
    const start = direction === 'out' ? 'scale(1.5)' : 'scale(.5)'
    frames = [{ transform: start, opacity: 0 }, { transform: 'scale(1)', opacity: 1 }]
  } else {
    const start = direction === 'r' ? 'inset(0 0 0 100%)' : direction === 'u' ? 'inset(100% 0 0 0)' : direction === 'd' ? 'inset(0 0 100% 0)' : 'inset(0 100% 0 0)'
    frames = [{ clipPath: start }, { clipPath: 'inset(0 0 0 0)' }]
  }
  const timing: KeyframeAnimationOptions = { duration: transition.durationMs, easing: effect === 'zoom' || effect === 'flythrough' || effect === 'warp' || effect === 'vortex' ? 'ease-in-out' : 'ease-out', fill: 'both', delay: animationDelay }
  const outgoing = underlay.value?.firstElementChild as HTMLElement | null
  let outgoingAnimation: Animation | undefined
  if (outgoingFrames && outgoing) outgoingAnimation = outgoing.animate(outgoingFrames, timing)
  if (effect === 'fade' && transition.throughBlack) {
    shell.value?.classList.add('is-fading-through-black')
    outgoing?.animate([{ opacity: 1, offset: 0 }, { opacity: 0, offset: 0.5 }, { opacity: 0, offset: 1 }], { ...timing, easing: 'linear' })
  }
  if (revealThroughBlack) shell.value?.classList.add('is-fading-through-black')
  let animation: Animation
  if (shredOutgoing && outgoing) {
    animation = outgoing.animate(frames, timing)
  } else if (shredOutgoing) {
    underlay.value?.classList.remove('is-shredding-out')
    underlay.value?.replaceChildren()
    return
  } else if (effect === 'pull' && outgoing) {
    const exits: Record<string, string> = {
      l: 'translateX(-100%)', r: 'translateX(100%)', u: 'translateY(-100%)', d: 'translateY(100%)',
      ld: 'translate(-100%, 100%)', lu: 'translate(-100%, -100%)', rd: 'translate(100%, 100%)', ru: 'translate(100%, -100%)',
    }
    const baseTransform = getComputedStyle(outgoing).transform
    const start = baseTransform === 'none' ? 'translate(0, 0)' : baseTransform
    animation = outgoing.animate([{ transform: start }, { transform: `${start} ${exits[direction] || exits.l}` }], timing)
  } else if (effect === 'pull') {
    underlay.value?.replaceChildren()
    return
  } else if (effect === 'split' && transition.direction === 'in' && outgoing) {
    const center = transition.orientation === 'horz' ? 'inset(50% 0)' : 'inset(0 50%)'
    animation = outgoing.animate([{ clipPath: 'inset(0)' }, { clipPath: center }], timing)
  } else if (effect === 'split' && transition.direction === 'in') {
    underlay.value?.replaceChildren()
    return
  } else {
    if (outgoingTransform && outgoing) {
      const baseTransform = getComputedStyle(outgoing).transform
      const start = baseTransform === 'none' ? 'translate(0, 0)' : baseTransform
      const end = `${start} ${outgoingTransform}`
      const keyframes = revealThroughBlack
        ? [{ transform: start, opacity: 1 }, { transform: end, opacity: 0, offset: 0.5 }, { transform: end, opacity: 0 }]
        : [{ transform: start }, { transform: end }]
      outgoing.animate(keyframes, timing)
    }
    animation = incoming.animate(frames, (effect === 'fade' || revealThroughBlack) && transition.throughBlack ? { ...timing, easing: 'linear' } : timing)
  }
  animation.onfinish = () => {
    animation.cancel()
      outgoingAnimation?.cancel()
    if (slide.value?.id === activeSlide.id) {
      underlay.value?.replaceChildren()
        underlay.value?.classList.remove('is-shredding-out', 'is-flythrough', 'is-warping-out', 'is-vortexing')
      shell.value?.classList.remove('is-fading-through-black')
    }
  }
  }, { flush: 'post' })
}

