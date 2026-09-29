import { nextTick, watch, type Ref } from 'vue'
import type { PptxSlide, PptxTransitionPlayback } from '~/utils/pptx'
import { combClipPath, dissolveClipPath, glitterTransitionClipPath, gridTransitionClipPath, plusTransitionClipPath, randomBarClipPath, rippleTransitionClipPath, shredTransitionClipPath, stripsClipPath, wedgeTransitionClipPath, wheelClipPath } from '~/utils/transitionMasks'

interface SlideTransitionOptions {
  slide: Readonly<Ref<PptxSlide | null>>
  shell: Ref<HTMLElement | null>
  incomingHost: Ref<HTMLElement | null>
  underlay: Ref<HTMLElement | null>
  flash: Ref<HTMLElement | null>
  playback: Readonly<Ref<PptxTransitionPlayback | null>>
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


export function useSlideTransition({ slide, shell, incomingHost, underlay, flash, playback }: SlideTransitionOptions) {
  watch(slide, (nextSlide, previousSlide) => {
    if (nextSlide?.id === previousSlide?.id) return
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
    underlay.value?.append(snapshot)
  }, { flush: 'sync' })

  watch(slide, async (activeSlide, previousSlide) => {
    if (!activeSlide || activeSlide.id === previousSlide?.id) return
    const currentPlayback = playback.value?.slideId === activeSlide.id ? playback.value : null
    const transition = currentPlayback?.transition
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
    stage.classList.add('is-morph-transition')
    if (transition.morphOption !== 'byObject') {
      const incomingAnimation = incoming.animate([{ opacity: 0 }, { opacity: 1 }], timing)
      const outgoingAnimation = outgoing.animate([{ opacity: 1 }, { opacity: 0 }], timing)
      animations.push(incomingAnimation, outgoingAnimation)
      incomingAnimation.onfinish = () => {
        animations.forEach(animation => animation.cancel())
        if (slide.value?.id === activeSlide.id) {
          underlay.value?.replaceChildren()
          stage.classList.remove('is-morph-transition')
        }
      }
      return
    }
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
      backgroundAnimation.onfinish = () => {
        animations.forEach(animation => animation.cancel())
        if (slide.value?.id === activeSlide.id) {
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
    const incomingBase = incoming.style.transform || 'none'
    const outgoingBase = outgoing.style.transform || 'none'
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
    const incomingAnimation = incoming.animate([{ transform: startTransform }, { transform: incomingBase }], timing)
    const outgoingAnimation = outgoing.animate([{ transform: outgoingBase }, { transform: endTransform }], timing)
    incomingAnimation.onfinish = () => {
      incomingAnimation.cancel()
      outgoingAnimation.cancel()
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
      clipPath: dissolveClipPath(width, height, progress),
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

