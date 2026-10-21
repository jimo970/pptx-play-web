import assert from 'node:assert/strict'
import { spawn } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync } from 'node:fs'
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { DOMParser } from '@xmldom/xmldom'

const fixture = process.env.DECKLINE_OFFICE_ANIMATION_FIXTURE || process.env.DECKLINE_OFFICE_MOTION_FIXTURE
const output = process.env.DECKLINE_OFFICE_ANIMATION_EVIDENCE
assert.ok(fixture && output, 'Set DECKLINE_OFFICE_ANIMATION_FIXTURE and DECKLINE_OFFICE_ANIMATION_EVIDENCE.')
const outputRoot = resolve(output)
if (existsSync(outputRoot)) assert.equal((await readdir(outputRoot)).length, 0, 'Use a new or empty evidence directory.')
await mkdir(outputRoot, { recursive: true })
const fixtureDirectory = dirname(fixture)
const metadataFile = existsSync(join(fixtureDirectory, 'fixture-metadata.json')) ? 'fixture-metadata.json' : 'metadata.json'
const metadata = JSON.parse(await readFile(join(fixtureDirectory, metadataFile), 'utf8'))
assert.ok(['Motion', 'Iterated', 'Word', 'CommonPresets', 'Emphasis', 'Transitions'].includes(metadata.Suite), 'Generate a supported suite with create-office-animation-fixture.ps1.')
const bounceOnly = process.env.DECKLINE_VERIFY_BOUNCE_ONLY === '1'
assert.ok(!bounceOnly || metadata.Suite === 'Motion', 'The Bounce-only verification requires a Motion-suite fixture.')
const requestedMotionPathTimes = process.env.DECKLINE_VERIFY_MOTION_PATH_TIMES_MS
  ?.split(',')
  .map(value => Number(value.trim()))
assert.ok(!requestedMotionPathTimes || requestedMotionPathTimes.length > 0 && requestedMotionPathTimes.every(value => Number.isInteger(value) && value >= 0), 'Motion path probe times must be non-negative integers separated by commas.')
const motionPathProbeTimes = requestedMotionPathTimes || [500, 1000, 1500, 2000]
const requestedTransitionTimes = process.env.DECKLINE_VERIFY_TRANSITION_TIMES_MS
  ?.split(',')
  .map(value => Number(value.trim()))
assert.ok(!requestedTransitionTimes || requestedTransitionTimes.length > 0 && requestedTransitionTimes.every(value => Number.isInteger(value) && value >= 0), 'Transition probe times must be non-negative integers separated by commas.')
const transitionProbeTimes = requestedTransitionTimes || [0]
const requestedCaseName = process.env.DECKLINE_VERIFY_CASE_NAME
const runtimeProcess = metadata.NativeProcesses?.find(process => process.Path?.toLowerCase().startsWith(metadata.ApplicationPath?.toLowerCase()))
const nativeRuntime = runtimeProcess ? {company: runtimeProcess.Company, productVersion: runtimeProcess.ProductVersion, executable: runtimeProcess.Path} : null
const isWpsRuntime = /kingsoft|wps/i.test([runtimeProcess?.Company, runtimeProcess?.Path].filter(Boolean).join(' '))
const fixtureBytes = await readFile(fixture)
const zip = await JSZip.loadAsync(fixtureBytes)
const parse = source => new DOMParser().parseFromString(source, 'text/xml')
const descendants = (node, name) => [...node.getElementsByTagName('*')].filter(element => element.localName === name)
const filterPoints = raw => raw.split(';').map(pair => pair.split(',').map(Number))
const clipPathPoints = path => {
  const values = path?.match(/-?\d+(?:\.\d+)?/g)?.map(Number) || []
  return Array.from({length: Math.floor(values.length / 2)}, (_, index) => [values[index * 2], values[index * 2 + 1]])
}
function pointInPolygon([x, y], points) {
  let inside = false
  for (let index = 0, previous = points.length - 1; index < points.length; previous = index++) {
    const [currentX, currentY] = points[index]
    const [previousX, previousY] = points[previous]
    if ((currentY > y) !== (previousY > y) && x < (previousX - currentX) * (y - currentY) / (previousY - currentY) + currentX) inside = !inside
  }
  return inside
}
function motionPathPosition(path, slideWidth, slideHeight, progress) {
  const tokens = path.match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?|[MLCZE]/gi) || []
  const numberPattern = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i
  const points = []
  const scale = ([x, y]) => [x * slideWidth, y * slideHeight]
  let cursor = 0
  let current = [0, 0]
  let pathStart
  let distance = 0
  const append = (point, snap = false) => {
    const [x, y] = point
    const previous = points.at(-1)
    if (previous && !snap) distance += Math.hypot(x - previous.x, y - previous.y)
    points.push({x, y, distance})
  }
  const readPoint = () => {
    const x = tokens[cursor++], y = tokens[cursor++]
    if (!numberPattern.test(x || '') || !numberPattern.test(y || '')) throw new Error(`Invalid native motion-path coordinate in ${path}`)
    return [Number(x), Number(y)]
  }
  while (cursor < tokens.length) {
    const action = tokens[cursor++]
    const normalized = action.toUpperCase()
    if (normalized === 'E') break
    if (normalized === 'M') {
      const coordinates = readPoint()
      current = action === normalized ? coordinates : [current[0] + coordinates[0], current[1] + coordinates[1]]
      pathStart = [...current]
      append(scale(current), true)
      continue
    }
    if (normalized === 'L') {
      const coordinates = readPoint()
      current = action === normalized ? coordinates : [current[0] + coordinates[0], current[1] + coordinates[1]]
      append(scale(current))
      continue
    }
    if (normalized === 'C') {
      const control1Input = readPoint(), control2Input = readPoint(), endInput = readPoint()
      const relative = action !== normalized
      const control1 = relative ? [current[0] + control1Input[0], current[1] + control1Input[1]] : control1Input
      const control2 = relative ? [current[0] + control2Input[0], current[1] + control2Input[1]] : control2Input
      const end = relative ? [current[0] + endInput[0], current[1] + endInput[1]] : endInput
      const from = scale(current), first = scale(control1), second = scale(control2), to = scale(end)
      for (let step = 1; step <= 4096; step++) {
        const t = step / 4096, inverse = 1 - t
        append([
          inverse ** 3 * from[0] + 3 * inverse ** 2 * t * first[0] + 3 * inverse * t ** 2 * second[0] + t ** 3 * to[0],
          inverse ** 3 * from[1] + 3 * inverse ** 2 * t * first[1] + 3 * inverse * t ** 2 * second[1] + t ** 3 * to[1]
        ])
      }
      current = end
      continue
    }
    if (normalized === 'Z') {
      if (!pathStart) throw new Error(`Native motion path closes before its first move: ${path}`)
      current = [...pathStart]
      append(scale(current))
      continue
    }
    throw new Error(`Unsupported native motion-path command ${action} in ${path}`)
  }
  assert.ok(points.length > 0 && distance > 0, `Native motion path must have measurable travel: ${path}`)
  const targetDistance = distance * progress
  const right = points.findIndex(point => point.distance >= targetDistance)
  if (right <= 0) return {x: points[0].x / slideWidth, y: points[0].y / slideHeight}
  const before = points[right - 1], after = points[right]
  const ratio = after.distance > before.distance ? (targetDistance - before.distance) / (after.distance - before.distance) : 0
  return {
    x: (before.x + (after.x - before.x) * ratio) / slideWidth,
    y: (before.y + (after.y - before.y) * ratio) / slideHeight
  }
}
function motionPathProgress(reference, progress) {
  const accelerationProgress = (curve, input) => {
    const acceleration = curve.acceleration || 0
    const deceleration = curve.deceleration || 0
    const area = 1 - (acceleration + deceleration) / 2
    if (acceleration && input < acceleration) return input ** 2 / (2 * acceleration * area)
    if (deceleration && input > 1 - deceleration) {
      const tail = input - (1 - deceleration)
      return (1 - deceleration - acceleration / 2 + tail - tail ** 2 / (2 * deceleration)) / area
    }
    return (acceleration / 2 + input - acceleration) / area
  }
  return reference.timingCurves.reduce((mapped, curve) => {
    const accelerated = accelerationProgress(curve, mapped)
    return curve.filter ? filteredProgress(curve.filter, accelerated) : accelerated
  }, progress)
}
let accent2Rgb
const themePath = Object.keys(zip.files).find(path => /^ppt\/theme\/theme\d+\.xml$/i.test(path))
if (themePath) {
  const theme = parse(await zip.file(themePath).async('string'))
  const accent2 = descendants(theme, 'accent2')[0]
  const color = accent2 && (descendants(accent2, 'srgbClr')[0]?.getAttribute('val') || descendants(accent2, 'sysClr')[0]?.getAttribute('lastClr'))
  if (color && /^[0-9a-f]{6}$/i.test(color)) accent2Rgb = color.match(/.{2}/g).map(channel => parseInt(channel, 16))
}
const presentation = parse(await zip.file('ppt/presentation.xml').async('string'))
const slideWidth = Number(descendants(presentation, 'sldSz')[0].getAttribute('cx'))
const slideHeight = Number(descendants(presentation, 'sldSz')[0].getAttribute('cy'))
const cases = new Map()
for (const [index, item] of metadata.Cases.entries()) {
  if (item.Supported === false) continue
  const doc = parse(await zip.file(`ppt/slides/slide${index + 1}.xml`).async('string'))
  const shapeName = descendants(doc, 'cNvPr').map(node => node.getAttribute('name')).find(Boolean)
  const shape = descendants(doc, 'sp').find(node => descendants(node, 'cNvPr').some(property => property.getAttribute('name') === shapeName))
  const offset = shape && descendants(shape, 'off')[0]
  const extent = shape && descendants(shape, 'ext')[0]
  assert.ok(offset, `${item.Name} must have a shape position in the native fixture`)
  const effectFilter = descendants(doc, 'animEffect')[0]?.getAttribute('filter')
  let motionPathReference
  if (item.Name.startsWith('Path')) {
    const motion = descendants(doc, 'animMotion')[0]
    const timing = motion && descendants(motion, 'cTn')[0]
    const parseTimingCurve = clock => ({
      acceleration: Number(clock.getAttribute('accel') || 0) / 100_000,
      deceleration: Number(clock.getAttribute('decel') || 0) / 100_000,
      filter: clock.getAttribute('tmFilter') ? filterPoints(clock.getAttribute('tmFilter')) : undefined
    })
    const parentTimingNodes = []
    for (let parent = motion?.parentElement; parent; parent = parent.parentElement) {
      if (parent.namespaceURI === 'http://schemas.openxmlformats.org/presentationml/2006/main' && parent.localName === 'cTn') parentTimingNodes.push(parent)
    }
    const timingCurves = [...parentTimingNodes.reverse(), ...(timing ? [timing] : [])]
      .filter(clock => ['accel', 'decel', 'tmFilter'].some(attribute => clock.hasAttribute(attribute)))
      .map(parseTimingCurve)
    const expectedPathEditMode = item.Relative === false ? 'fixed' : 'relative'
    assert.ok(motion?.getAttribute('path'), `${item.Name} must export a native p:animMotion path`)
    assert.equal(motion.getAttribute('origin'), 'layout', `${item.Name} must retain the WPS layout origin`)
    assert.equal(motion.getAttribute('pathEditMode'), expectedPathEditMode, `${item.Name} must preserve its WPS relative/absolute path mode`)
    assert.equal(Number(timing?.getAttribute('dur')), 2000, `${item.Name} must preserve its two-second native path duration`)
    assert.ok(extent, `${item.Name} must retain its native shape dimensions`)
    motionPathReference = {
      path: motion.getAttribute('path'),
      pathEditMode: expectedPathEditMode,
      durationMs: Number(timing.getAttribute('dur')),
      timingCurves,
      shapeCenter: {
        x: (Number(offset.getAttribute('x')) + Number(extent.getAttribute('cx')) / 2) / slideWidth,
        y: (Number(offset.getAttribute('y')) + Number(extent.getAttribute('cy')) / 2) / slideHeight
      }
    }
  }
  let peekReference
  let scaleTiming
  if (item.Name === 'FlashBulb') {
    const animation = descendants(doc, 'animScale')[0]
    const clock = animation && descendants(animation, 'cTn')[0]
    const relativeScale = animation && descendants(animation, 'by')[0]
    const rawAutoReverse = clock?.getAttribute('autoRev')
    scaleTiming = {
      durationMs: Number(clock?.getAttribute('dur')),
      autoReverse: rawAutoReverse === '1' || rawAutoReverse === 'true',
      scaleX: Number(relativeScale?.getAttribute('x')) / 100_000,
      scaleY: Number(relativeScale?.getAttribute('y')) / 100_000,
    }
    assert.ok(scaleTiming.durationMs > 0 && scaleTiming.autoReverse
      && Number.isFinite(scaleTiming.scaleX) && Number.isFinite(scaleTiming.scaleY),
    'FlashBulb must preserve finite relative scale values and an auto-reversing scale clock')
  }
  if (isWpsRuntime && item.Name === 'Peek') {
    const motion = descendants(doc, 'anim').find(animation => descendants(animation, 'attrName')[0]?.textContent === 'ppt_y')
    assert.ok(motion, 'WPS Peek must export a vertical shape-motion track')
    const keyframes = descendants(motion, 'tav').map(point => ({
      time: Number(point.getAttribute('tm')),
      formula: descendants(point, 'strVal')[0]?.getAttribute('val')
    }))
    const startOffset = /^#ppt_y\+#ppt_h\*([0-9]+(?:\.[0-9]+)?)$/.exec(keyframes[0]?.formula || '')
    const motionDurationMs = Number(descendants(motion, 'cTn')[0]?.getAttribute('dur'))
    const wipe = descendants(doc, 'animEffect').find(effect => effect.getAttribute('filter') === 'wipe(up)')
    const wipeDurationMs = Number(descendants(wipe, 'cTn')[0]?.getAttribute('dur'))
    const preset = descendants(doc, 'cTn').find(clock => clock.getAttribute('presetID') === '12')
    assert.deepEqual(keyframes.map(point => point.time), [0, 100000], 'WPS Peek must preserve its linear start and end points')
    assert.equal(keyframes[1].formula, '#ppt_y', 'WPS Peek must return the shape to its saved vertical position')
    assert.ok(startOffset, `WPS Peek must express its entrance distance in terms of the saved shape height: ${keyframes[0]?.formula}`)
    assert.equal(effectFilter, 'wipe(up)', 'WPS Peek must pair its motion track with an upward wipe')
    assert.ok(motionDurationMs > 0 && motionDurationMs === wipeDurationMs, 'WPS Peek motion and wipe tracks must share one finite duration')
    assert.equal(preset?.getAttribute('presetSubtype'), '4', 'WPS Peek must preserve its upward preset subtype')
    peekReference = {durationMs: motionDurationMs, startOffsetRatio: Number(startOffset[1]), wipeFilter: effectFilter}
  }
  const randomEffectPlaceholder = item.Name === 'RandomEffects' && descendants(doc, 'cTn').some(owner =>
    owner.getAttribute('presetID') === '24' && owner.getAttribute('presetClass') === 'entr'
      && descendants(owner, 'anim').some(animation => !descendants(animation, 'attrName').some(node => node.textContent?.trim()) && animation.getAttribute('to') === '')
      && !descendants(owner, 'animEffect').length)
  const iterated = descendants(doc, 'iterate').length > 0
  const iterationType = descendants(doc, 'iterate')[0]?.getAttribute('type')
  const expectedIterated = item.ByWord === true || item.ByCharacter !== false
  if (!['CommonPresets', 'Emphasis', 'Transitions'].includes(metadata.Suite)) assert.equal(iterated, expectedIterated, `${item.Name} must preserve the requested text-unit animation`)
  if (iterated) assert.equal(iterationType, item.ByWord ? 'wd' : 'lt', `${item.Name} must preserve its character or word iteration type`)
  let rotationStartDegrees
  if (item.Name === 'Swish') {
    const rotationSet = descendants(doc, 'set').find(set => descendants(set, 'attrName')[0]?.textContent === 'style.rotation')
    const rotationTo = rotationSet && descendants(rotationSet, 'to')[0]
    rotationStartDegrees = Number(rotationTo && descendants(rotationTo, 'strVal')[0]?.getAttribute('val'))
    assert.equal(rotationStartDegrees, -45, 'Swish must retain its -45-degree p:set starting rotation')
  }
  let rotationKeyframes
  if (item.Name === 'GrowAndTurn') {
    const rotationAnimation = descendants(doc, 'anim').find(animation => descendants(animation, 'attrName')[0]?.textContent === 'style.rotation')
    rotationKeyframes = rotationAnimation && descendants(rotationAnimation, 'tav').map(point => Number(descendants(point, 'fltVal')[0]?.getAttribute('val')))
    assert.deepEqual(rotationKeyframes, [90, 0], 'GrowAndTurn must retain its 90-to-0 degree rotation track')
  }
  let fillColorTiming
  if (['Blast', 'VerticalGrow'].includes(item.Name)) {
    const setValue = property => {
      const set = descendants(doc, 'set').find(node => descendants(node, 'attrName')[0]?.textContent === property)
      return set && descendants(set, 'strVal')[0]?.getAttribute('val')
    }
    const fillColorAnimation = descendants(doc, 'animClr').find(animation => descendants(animation, 'attrName')[0]?.textContent?.toLowerCase() === 'fillcolor')
    const fillColorClock = fillColorAnimation && descendants(fillColorAnimation, 'cTn')[0]
    const fillColorDelayMs = Number((fillColorClock && descendants(fillColorClock, 'cond')[0]?.getAttribute('delay')) || 0)
    fillColorTiming = {durationMs: Number(fillColorClock?.getAttribute('dur')), delayMs: fillColorDelayMs}
    if (isWpsRuntime) {
      assert.equal(setValue('fill.type'), 'solid', `${item.Name} must activate a solid fill type`)
      assert.equal(setValue('fill.on'), 'true', `${item.Name} must enable the whole-object fill`)
    }
    assert.ok(fillColorAnimation && descendants(fillColorAnimation, 'from').length === 0, `${item.Name} must preserve its to-only fill-color track`)
    assert.equal(descendants(fillColorAnimation, 'schemeClr').find(node => node.parentElement?.localName === 'to')?.getAttribute('val'), 'accent2', `${item.Name} must retain its accent2 fill destination`)
    assert.ok(fillColorTiming.durationMs > 0 && fillColorTiming.delayMs >= 0, `${item.Name} must preserve the fill-color timing`)
  }
  if (item.Exit) assert.ok(descendants(doc, 'cTn').some(node => node.getAttribute('presetClass') === 'exit'), `${item.Name} must retain its exit class`)
  let bounceReference
  let effectDurationMs
  let iterationIntervalMs
  let childStyleDelayMs
  if (iterated && ['Iterated', 'Word'].includes(metadata.Suite)) {
    // PowerPoint may place a 1ms cTn on the click group before the effect's own clock.
    // Derive the reference duration from the actual visual behavior, not an ancestor clock.
    const textUnitAnimation = ['animEffect', 'animClr', 'animScale', 'animRot', 'animMotion', 'anim']
      .map(name => descendants(doc, name)[0])
      .find(Boolean)
    const timing = textUnitAnimation && descendants(textUnitAnimation, 'cTn').find(node =>
      node.hasAttribute('dur') && Number.isFinite(Number(node.getAttribute('dur'))) && Number(node.getAttribute('dur')) > 0)
    const interval = descendants(doc, 'iterate').flatMap(node => descendants(node, 'tmPct'))[0]
    effectDurationMs = Number(timing?.getAttribute('dur'))
    iterationIntervalMs = effectDurationMs * Number(interval?.getAttribute('val')) / 100000
    assert.ok(effectDurationMs > 0 && iterationIntervalMs > 0, `${item.Name} must preserve native effect duration and text iteration interval`)
    const preset = descendants(doc, 'cTn').find(node => node.getAttribute('nodeType') === 'clickEffect')
    assert.equal(Number(preset?.getAttribute('presetID')), item.PresetId ?? item.Id, `${item.Name} must preserve the native preset ID`)
    assert.equal(preset?.getAttribute('presetClass'), item.PresetClass || 'entr', `${item.Name} must preserve the native preset class`)
    if (item.Name === 'ColorByWord') {
      const animation = descendants(doc, 'animClr')[0]
      const targetProperty = descendants(animation, 'attrName')[0]?.textContent
      const destination = descendants(animation, 'schemeClr').find(node => node.parentElement?.localName === 'to')
      assert.ok(animation, 'ColorByWord must be a native p:animClr effect')
      assert.equal(targetProperty, 'style.color', 'ColorByWord must target text color')
      assert.equal(descendants(animation, 'from').length, 0, 'This native reference uses the documented omitted-from color behavior')
      assert.equal(destination?.getAttribute('val'), 'accent2', 'ColorByWord must retain the native theme color destination')
      assert.ok(accent2Rgb, 'The native presentation theme must define accent2')
    }
  }
  if (item.Name === 'GrowWithColor') {
    const preset = descendants(doc, 'cTn').find(node => node.getAttribute('nodeType') === 'clickEffect' && node.getAttribute('presetID') === '28')
    const iterate = preset && descendants(preset, 'iterate')[0]
    const colorTrack = preset && descendants(preset, 'animClr').find(node => descendants(node, 'attrName')[0]?.textContent === 'style.color')
    const fillTrack = preset && descendants(preset, 'animClr').find(node => descendants(node, 'attrName')[0]?.textContent === 'fillcolor')
    const fontSizeTrack = preset && descendants(preset, 'anim').find(node => descendants(node, 'attrName')[0]?.textContent === 'style.fontSize')
    effectDurationMs = Number(fontSizeTrack && descendants(fontSizeTrack, 'cTn')[0]?.getAttribute('dur'))
    const colorDurationMs = Number(colorTrack && descendants(colorTrack, 'cTn')[0]?.getAttribute('dur'))
    const fillDurationMs = Number(fillTrack && descendants(fillTrack, 'cTn')[0]?.getAttribute('dur'))
    childStyleDelayMs = isWpsRuntime ? fillDurationMs : 0
    const interval = iterate && descendants(iterate, 'tmPct')[0]
    iterationIntervalMs = effectDurationMs * Number(interval?.getAttribute('val')) / 100000
    assert.ok(preset && colorTrack && fillTrack && fontSizeTrack && effectDurationMs > 0
      && colorDurationMs === effectDurationMs && fillDurationMs > 0 && iterationIntervalMs > 0,
      'GrowWithColor must retain its childStyle, fill-color, and iterated timing tracks')
  }
  if (metadata.Suite === 'Motion' && item.Name === 'Bounce') {
    const coordinates = descendants(doc, 'anim').map(animation => {
      const clock = descendants(animation, 'cTn')[0]
      const values = descendants(animation, 'tav')
      const formula = values[0].getAttribute('fmla')
      const property = descendants(animation, 'attrName')[0].textContent
      return {
        property, duration: Number(clock.getAttribute('dur')), delay: Number(descendants(clock, 'cond')[0]?.getAttribute('delay') || 0),
        filter: filterPoints(clock.getAttribute('tmFilter')),
        input: values.map(value => Number(descendants(value, 'fltVal')[0]?.getAttribute('val'))), formula
      }
    })
    const horizontal = coordinates.find(track => track.property === 'ppt_x')
    const vertical = coordinates.filter(track => track.property === 'ppt_y')
    assert.equal(vertical.length, 4, 'The native Bounce reference must preserve its four vertical legs')
    assert.ok(horizontal?.duration > 0, 'The native Bounce reference must preserve its horizontal leg')
    for (const track of vertical) {
      assert.match(track.formula, /^#ppt_y-sin\(pi\*\$\)\/(3|9|27|81)$/)
      track.divisor = Number(track.formula.split('/').at(-1))
      assert.ok(track.input.every(Number.isFinite) && track.duration > 0, 'Bounce reference inputs must be finite')
    }
    const scales = descendants(doc, 'animScale').map(animation => {
      const clock = descendants(animation, 'cTn')[0]
      const destination = descendants(animation, 'to')[0]
      const additive = descendants(animation, 'cBhvr')[0]?.getAttribute('additive')
      assert.equal(descendants(animation, 'from').length, 0, 'This reference uses documented to-only scale behavior')
      assert.ok(additive === null || additive === 'base', 'The native Bounce scale tracks must use base-additive composition')
      return {
        duration: Number(clock.getAttribute('dur')), delay: Number(descendants(clock, 'cond')[0]?.getAttribute('delay') || 0),
        deceleration: Number(clock.getAttribute('decel') || 0) / 100000,
        to: ['x', 'y'].map(axis => Number(destination.getAttribute(axis)) / 100000),
        fromCurrent: true
      }
    })
    assert.equal(scales.length, 8, 'The native Bounce reference must preserve all eight landing scale tracks')
    assert.ok(scales.every(track => track.fromCurrent), 'Every native base-additive to-only Bounce scale track must start from the currently rendered shape scale')
    const preset = descendants(doc, 'cTn').find(node => node.getAttribute('nodeType') === 'clickEffect')
    const iterationNode = descendants(preset, 'iterate')[0]
    const intervalNode = descendants(iterationNode, 'tmPct')[0]
    const iterationDurationMs = descendants(preset, 'cTn').reduce((duration, node) => {
      const nodeDuration = Number(node.getAttribute('dur'))
      const nodeDelay = Number(descendants(node, 'cond')[0]?.getAttribute('delay') || 0)
      return Number.isFinite(nodeDuration) && nodeDuration >= 0 ? Math.max(duration, nodeDelay + nodeDuration) : duration
    }, 0)
    const rawInterval = intervalNode?.getAttribute('val')?.trim() || ''
    const intervalPercentage = rawInterval.endsWith('%') ? Number(rawInterval.slice(0, -1)) / 100 : /^\d+$/.test(rawInterval) ? Number(rawInterval) / 100_000 : Number.NaN
    const iterationIntervalMs = iterationDurationMs * intervalPercentage
    assert.equal(iterationDurationMs, 2000, 'Bounce iteration duration must include the delayed landing-scale tracks through 2,000ms')
    assert.equal(iterationIntervalMs, 200, 'Bounce tmPct must stagger letters by 200ms for the complete behavior span')
    bounceReference = {horizontal, vertical, scales, iterationIntervalMs}
  }
  const fontSizeNode = shape && descendants(shape, 'rPr').find(node => node.hasAttribute('sz'))
  const staticTextColorNode = shape && descendants(shape, 'rPr').flatMap(node => descendants(node, 'srgbClr'))[0]
  const staticTextColorHex = staticTextColorNode?.getAttribute('val')
  const staticTextColorRgb = staticTextColorHex && /^[0-9a-f]{6}$/i.test(staticTextColorHex) ? staticTextColorHex.match(/.{2}/g).map(channel => parseInt(channel, 16)) : undefined
  const fontFamilySet = descendants(doc, 'set').find(node => descendants(node, 'attrName')[0]?.textContent === 'style.fontFamily')
  const fontFamilyTo = fontFamilySet && descendants(fontFamilySet, 'strVal')[0]?.getAttribute('val')
  const fontStyleSet = descendants(doc, 'set').find(node => descendants(node, 'attrName')[0]?.textContent === 'style.fontStyle')
  const fontStyleTo = fontStyleSet && descendants(fontStyleSet, 'strVal')[0]?.getAttribute('val')
  const fontWeightSet = descendants(doc, 'set').find(node => descendants(node, 'attrName')[0]?.textContent === 'style.fontWeight')
  const fontWeightTo = fontWeightSet && descendants(fontWeightSet, 'strVal')[0]?.getAttribute('val')
  cases.set(item.Name, {...item, slide: index + 1, shapeName, y: Number(offset.getAttribute('y')), effectFilter, iterated, randomEffectPlaceholder, bounceReference, motionPathReference, effectDurationMs, iterationIntervalMs, childStyleDelayMs, rotationStartDegrees, rotationKeyframes, fillColorTiming, peekReference, scaleTiming, fontSizePt: Number(fontSizeNode?.getAttribute('sz')) / 100, staticTextColorRgb, fontFamilyTo, fontStyleTo, fontWeightTo})
}
const motionProbes = [
  ['Fly', 500], ['Bounce', 500], ['Bounce', 658], ['Bounce', 750], ['Bounce', 1000],
  ['Bounce', 1320], ['Bounce', 1400], ['Bounce', 1650], ['Bounce', 1720],
  ['Bounce', 1816], ['Bounce', 1900], ['Bounce', 2700],
  ['Spin', 500], ['GrowShrink', 500], ['GrowAndTurn', 500], ['Blast', 500], ['Blast', 2700], ['RandomEffects', 500], ['RandomEffects', 2700], ['VerticalGrow', 500], ['VerticalGrow', 4000], ['Swivel', 1200], ['Swish', 500], ['FadedSwivel', 1200],
  ['FadeExit', 500], ['ZoomExit', 500], ['FlyExit', 500],
  ...[...cases.values()]
    .filter(item => item.Supported !== false && item.Name.startsWith('Path'))
    .flatMap(item => motionPathProbeTimes.map(targetMs => [item.Name, targetMs])),
  ['FadeExit', 2700], ['ZoomExit', 2700], ['FlyExit', 2700]
]
const iteratedNames = ['Fade', 'Zoom', 'Checkerboard', 'Blinds', 'Dissolve', 'Wheel', 'Box', 'RandomBars', 'Strips', 'Wipe']
const iteratedProbes = [...iteratedNames.map(name => [name, 500]), ...iteratedNames.map(name => [name, 2700])]
const wordNames = ['FadeByWord', 'WipeByWord', 'RandomBarsByWord', 'StripsByWord', 'ColorByWord']
const wordProbes = [...wordNames.map(name => [name, 500]), ...wordNames.map(name => [name, 2700])]
const commonPresetNames = metadata.Cases.map(item => item.Name)
const commonPresetProbes = [
  ...commonPresetNames.map(name => [name, 500]),
  ...commonPresetNames.map(name => [name, 7000]),
  ['Swish', 2700], ['Swish', 6000], ['VerticalGrow', 2700], ['VerticalGrow', 4000]
]
const transitionProbes = metadata.Cases
  .map((item, index) => ({item, slide: index + 1}))
  .filter(({item}) => item.Supported !== false)
  .flatMap(({item, slide}) => slide === 1
    ? [[item.Name, 0, slide]]
    : transitionProbeTimes.map(targetMs => [item.Name, targetMs, slide]))
const supportedCases = metadata.Cases.filter(item => item.Supported !== false)
const emphasisNames = supportedCases.map(item => item.Name)
const emphasisProbes = [
  ...emphasisNames.map(name => [name, 500]),
  ...emphasisNames.map(name => [name, 2500]),
  ...(emphasisNames.includes('BoldFlash') ? [['BoldFlash', 1100]] : []),
  ...(emphasisNames.includes('FlashBulb') ? [['FlashBulb', 2000], ['FlashBulb', 4500]] : []),
]
const suiteProbes = metadata.Suite === 'Transitions' ? transitionProbes : metadata.Suite === 'Iterated' ? iteratedProbes : metadata.Suite === 'Word' ? wordProbes : metadata.Suite === 'CommonPresets' ? commonPresetProbes : metadata.Suite === 'Emphasis' ? emphasisProbes : motionProbes
const selectedSuiteProbes = requestedCaseName ? suiteProbes.filter(([name]) => name === requestedCaseName) : suiteProbes
assert.ok(!requestedCaseName || selectedSuiteProbes.length > 0, `No ${metadata.Suite} probes matched DECKLINE_VERIFY_CASE_NAME=${requestedCaseName}.`)
const probes = bounceOnly ? selectedSuiteProbes.filter(([name]) => name === 'Bounce') : selectedSuiteProbes
const report = {suite: metadata.Suite, fixture: resolve(fixture), fixtureSha256: createHash('sha256').update(fixtureBytes).digest('hex'), application: metadata.ApplicationPath, nativeRuntime, runtimeMode: process.env.DECKLINE_RUNTIME_MODE || 'UNSPECIFIED', clockMode: 'controlled', probes: [], unsupportedNativeCases: metadata.Cases.filter(item => item.Supported === false).map(item => ({name: item.Name, effectId: item.Id ?? item.EntryEffect, reason: item.SupportError})), result: 'RUNNING', limitations: [metadata.Suite === 'Transitions' ? 'Transition samples seek WAAPI currentTime after explicit navigation from the preceding slide; they verify browser keyframe interpolation, not native Office/WPS frame cadence or native pixel parity.' : 'Browser playback of native-exported OOXML is checked with a controlled performance.now clock; this does not establish real-time frame cadence or native Office/WPS animation-frame pixel parity.']}
const projectRoot = fileURLToPath(new URL('../', import.meta.url))
const captureEnvironment = Object.fromEntries(Object.entries(process.env).filter(([name]) => !name.startsWith('DECKLINE_VERIFY_')))
const randomEffectMidStates = new Map()

async function capture(name, elapsedMs, item) {
  const imagePath = join(outputRoot, `${name.toLowerCase()}-${elapsedMs}.png`)
  const isTransition = metadata.Suite === 'Transitions'
  const transitionStartsOnClick = isTransition && item.slide > 1
  const captureSlideNumber = transitionStartsOnClick ? item.slide - 1 : item.slide
  const animationClock = 'controlled'
  const attempts = []
  let result
  for (let attempt = 0; attempt < 3; attempt++) {
    result = await new Promise((resolveResult, reject) => {
      const child = spawn(process.execPath, [join(projectRoot, 'scripts/verify-animation.mjs')], {
        cwd: projectRoot,
        env: {...captureEnvironment, DECKLINE_INPUT_PPTX: resolve(fixture), DECKLINE_CAPTURE_ONLY: '1', DECKLINE_CAPTURE_SLIDE_NUMBER: String(captureSlideNumber), DECKLINE_CAPTURE_ANIMATION_MS: String(elapsedMs), DECKLINE_CAPTURE_ANIMATION_CLOCK: animationClock, ...(transitionStartsOnClick ? {DECKLINE_CAPTURE_START_ANIMATION: '1', DECKLINE_CAPTURE_ADVANCE_SLIDE: '1', DECKLINE_CAPTURE_WAIT_FOR_TRANSITION: '1', DECKLINE_CAPTURE_SEEK_TRANSITION_MS: String(elapsedMs)} : {}), DECKLINE_CAPTURE_SLIDE_PATH: imagePath},
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe']
      })
      let log = ''
      child.stdout.on('data', chunk => { log += chunk })
      child.stderr.on('data', chunk => { log += chunk })
      child.once('error', reject)
      child.once('exit', code => resolveResult({code, log}))
    })
    attempts.push(result.log)
    const transientBrowserStartupFailure = /Browser state: \{"url":"about:blank"|The headless browser did not expose the player page\./.test(result.log)
    if (result.code === 0 || !transientBrowserStartupFailure || attempt === 2) break
    await new Promise(resolveDelay => setTimeout(resolveDelay, 250))
  }
  await writeFile(`${imagePath}.log`, attempts.map((log, index) => `Capture attempt ${index + 1}:\n${log}`).join('\n'))
  assert.equal(result.code, 0, `Browser capture failed for ${name}: ${result.log.slice(-1600)}`)
  return JSON.parse(await readFile(`${imagePath}.json`, 'utf8'))
}

const close = (actual, expected, tolerance, label) => assert.ok(Math.abs(actual - expected) <= tolerance, `${label}: ${actual} != ${expected} (tolerance ${tolerance})`)
const rgbChannels = (value, label) => {
  const match = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(value || '')
  assert.ok(match, `${label} must expose a rendered RGB color: ${value}`)
  return match.slice(1).map(Number)
}
const expectRgb = (actual, expected, label) => {
  rgbChannels(actual, label).forEach((channel, index) => close(channel, expected[index], 1, `${label} channel ${index + 1}`))
}
const effectSummary = runs => runs.map(run => {
  const value = run.maskImage !== 'none' ? run.maskImage : run.clipPath
  return {text: run.text, visibility: run.visibility, opacity: run.opacity, color: run.color, scaleX: run.scaleX, scaleY: run.scaleY, clipPath: run.clipPath, effect: value === 'none' ? 'none' : createHash('sha256').update(value).digest('hex').slice(0, 10)}
})
// Interpolate the exported controls independently of the player's clock/formula helpers.
function filteredProgress(filter, progress) {
  if (progress <= 0) return 0
  if (progress >= 1) return 1
  const right = filter.findIndex(point => point[0] >= progress)
  if (right <= 0) return filter[Math.max(0, right)][1]
  const before = filter[right - 1], after = filter[right]
  return before[1] + (after[1] - before[1]) * (progress - before[0]) / (after[0] - before[0])
}

function bounceAt(reference, elapsedMs) {
  const leg = reference.vertical.filter(track => track.delay <= elapsedMs).at(-1)
  const progress = filteredProgress(leg.filter, (elapsedMs - leg.delay) / leg.duration)
  const input = leg.input[0] + (leg.input[1] - leg.input[0]) * progress
  let scale = [1, 1]
  for (const track of reference.scales) {
    if (track.delay > elapsedMs) break
    let scaleProgress = Math.min(1, Math.max(0, (elapsedMs - track.delay) / track.duration))
    if (track.deceleration) {
      const tail = Math.max(0, scaleProgress - 1 + track.deceleration)
      scaleProgress = (scaleProgress - tail * tail / (2 * track.deceleration)) / (1 - track.deceleration / 2)
    }
    const from = track.fromCurrent ? scale : [0, 0]
    scale = track.to.map((value, axis) => from[axis] + (value - from[axis]) * scaleProgress)
  }
  return {
    x: -0.25 * (1 - filteredProgress(reference.horizontal.filter, elapsedMs / reference.horizontal.duration)),
    y: -Math.sin(Math.PI * input) / leg.divisor,
    scale
  }
}

try {
  for (const [name, targetMs, transitionSlide] of probes) {
    const item = cases.get(name)
    assert.ok(item, `Missing native case ${name}`)
    const captureResult = await capture(name, targetMs, item)
    const frame = captureResult.frameRect
    const state = captureResult.captureAnimation
    if (metadata.Suite === 'Transitions' && transitionSlide > 1) {
      const elementNames = state.elements.map(element => element.name)
      assert.ok(elementNames.includes(item.shapeName), `${name} must sample the target slide content: ${JSON.stringify({slide: captureResult.slide, expected: item.shapeName, actual: elementNames})}`)
    }
    assert.equal(state.clockMode, 'controlled', 'Short landing tracks require the controlled capture clock')
    close(state.elapsedMs, targetMs, 0.001, 'The rendered checkpoint must use the requested clock value')
    if (metadata.Suite === 'Transitions') {
      const transition = captureResult.transition
      const effectForElement = {
        fade: 'fade', dissolve: 'dissolve', wipe: 'wipe', zoom: 'zoom', checker: 'checker', cover: 'cover', pull: 'pull', push: 'push',
        split: 'split', circle: 'circle', diamond: 'diamond', plus: 'plus', wedge: 'wedge', blinds: 'blinds', comb: 'comb', randomBar: 'randomBar',
        wheel: 'wheel', strips: 'strips', doors: 'doors', window: 'window', honeycomb: 'honeycomb', flythrough: 'flythrough', conveyor: 'conveyor'
      }
      const expectedEffect = effectForElement[item.ExportedEffect]
      assert.ok(expectedEffect, `${name} has an unrecognized exported transition element: ${item.ExportedEffect}`)
      assert.equal(transition?.effect, expectedEffect, `${name} must parse to the exported ${expectedEffect} transition`)
      const transitionXml = Array.isArray(item.ExportedTransitionXml) ? item.ExportedTransitionXml[0] : item.ExportedTransitionXml
      assert.ok(transitionXml, `${name} must retain exported transition XML in fixture metadata`)
      const transitionNode = parse(transitionXml).documentElement
      const rawDuration = transitionNode.getAttributeNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'dur')
        || transitionNode.getAttribute('p14:dur')
      const explicitDuration = /^(\d+(?:\.\d+)?)(ms|s)?$/i.exec(rawDuration || '')
      const speed = transitionNode.getAttribute('spd') || 'fast'
      const expectedDurationMs = explicitDuration
        ? Number(explicitDuration[1]) * (explicitDuration[2]?.toLowerCase() === 's' ? 1000 : 1)
        : ({fast: 500, med: 750, slow: 1000}[speed] || 500)
      if (transitionSlide === 1) {
        report.probes.push({name, slide: transitionSlide, effect: transition.effect, expectedDurationMs, runtime: 'initial slide parsed; no incoming transition is played on startup', result: 'PASS', evidence: `${name.toLowerCase()}-0.png.json`})
        console.log(`${name} parsed on the initial slide; startup has no incoming transition.`)
        await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
        continue
      }
      const incoming = transition?.incoming
      const outgoing = transition?.outgoing
      const durationMs = incoming?.duration ?? outgoing?.duration
      assert.equal(durationMs, expectedDurationMs, `${name} must preserve the exported transition duration`)
      const hasEffectTrack = [incoming?.keyframes, incoming?.opacityKeyframes, incoming?.clipKeyframes, outgoing?.keyframes, outgoing?.opacityKeyframes, outgoing?.clipKeyframes]
        .some(keyframes => Array.isArray(keyframes) && keyframes.length > 1)
      assert.ok(hasEffectTrack, `${name} must create a live incoming or outgoing Web Animation track`)
      assert.ok(!captureResult.warnings.some(warning => /falls back to fade|unsupported .* transition|unsupported transition/i.test(warning)), `${name} must not degrade to a fallback transition`)
      const currentTimeMs = incoming?.currentTime ?? outgoing?.currentTime
      assert.ok(Number.isFinite(currentTimeMs), `${name} must expose its live transition animation time`)
      if (targetMs > 0 && targetMs < durationMs) close(currentTimeMs, targetMs, 1, `${name} sought WAAPI currentTime`)
      if (name === 'StripsDownLeft' && targetMs > 0 && targetMs < durationMs) {
        assert.ok(incoming?.clipPath && incoming.clipPath !== incoming.clipKeyframes?.[0],
          `StripsDownLeft must render an intermediate strip mask at ${targetMs}ms: ${JSON.stringify(incoming)}`)
      }
      report.probes.push({name, slide: transitionSlide, targetMs, effect: transition.effect, durationMs, currentTimeMs, clipPath: incoming?.clipPath, result: 'PASS', evidence: `${name.toLowerCase()}-${targetMs}.png.json`})
      console.log(`${name} on slide ${transitionSlide} passed at ${Math.round(currentTimeMs)}ms of ${durationMs}ms.`)
      await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
      continue
    }
    const target = item.iterated ? state.runs[0] : state.elements.find(element => element.name === item.shapeName)
    assert.ok(target && frame?.height > 0, `Missing rendered target for ${name}`)
    let motionPathCheck
    if (item.motionPathReference) {
      const linearProgress = targetMs / item.motionPathReference.durationMs
      const progress = motionPathProgress(item.motionPathReference, linearProgress)
      const expected = motionPathPosition(item.motionPathReference.path, slideWidth, slideHeight, progress)
      const expectedOffset = item.motionPathReference.pathEditMode === 'fixed'
        ? [expected.x - item.motionPathReference.shapeCenter.x, expected.y - item.motionPathReference.shapeCenter.y]
        : [expected.x, expected.y]
      const expectedTranslation = [expectedOffset[0] * frame.width, expectedOffset[1] * frame.height]
      const actualTranslation = [target.matrix[4], target.matrix[5]]
      close(actualTranslation[0], expectedTranslation[0], 1.5, `${name} WPS path horizontal position at ${targetMs}ms`)
      close(actualTranslation[1], expectedTranslation[1], 1.5, `${name} WPS path vertical position at ${targetMs}ms`)
      assert.ok(!captureResult.warnings.some(warning => /motion path.*(skipped|unsupported)|unsupported.*motion path/i.test(warning)), `${name} must render the exported WPS path without fallback warnings`)
      motionPathCheck = {pathEditMode: item.motionPathReference.pathEditMode, linearProgress, effectiveProgress: progress, expectedTranslationPx: expectedTranslation, actualTranslationPx: actualTranslation}
    }
    if (metadata.Suite === 'CommonPresets' && name === 'Wedge') {
      assert.equal(item.effectFilter, 'wedge', 'The native WPS Wedge preset must export the wedge transition filter')
      assert.ok(!captureResult.warnings.some(warning => /unsupported object-effect filter|approximated as a fade/.test(warning)), 'The native WPS Wedge preset must render without an unsupported-filter fallback')
      const points = clipPathPoints(target.clipPath)
      assert.ok(points.length >= 4, `The native WPS Wedge preset must produce a polygon clip path: ${JSON.stringify(target)}`)
      if (targetMs < 2000) {
        const [center, firstRadialPoint, lastRadialPoint] = [points[0], points[1], points.at(-1)]
        const angle = ([x, y]) => (Math.atan2(y - center[1], x - center[0]) * 180 / Math.PI + 360) % 360
        const startAngle = angle(firstRadialPoint)
        const endAngle = angle(lastRadialPoint)
        const sweep = (endAngle - startAngle + 360) % 360
        const middleAngle = (startAngle + sweep / 2) % 360
        const topOriginError = Math.min(Math.abs(middleAngle - 270), 360 - Math.abs(middleAngle - 270))
        close(center[0], target.rect.width / 2, 2, 'Wedge mask horizontal origin')
        close(center[1], target.rect.height / 2, 2, 'Wedge mask vertical origin')
        assert.ok(sweep > 0 && sweep < 180 && topOriginError < 3,
          `The WPS Wedge mask must open symmetrically from 12 o’clock: ${JSON.stringify({startAngle, endAngle, sweep, middleAngle, topOriginError, clipPath: target.clipPath})}`)
      } else {
        const corners = [[0, 0], [target.rect.width, 0], [target.rect.width, target.rect.height], [0, target.rect.height]]
        assert.ok(corners.every(corner => pointInPolygon(corner, points)),
          `The completed WPS Wedge mask must reveal every shape corner: ${JSON.stringify({corners, points, clipPath: target.clipPath})}`)
      }
    }
    if (metadata.Suite === 'Emphasis') {
      const run = state.runs[0]
      assert.ok(run && target.visibility === 'visible', `${name} must retain its text during emphasis`)
      if (name === 'Shimmer' && targetMs === 500) {
        const runs = state.runs
        const scalesX = runs.map(character => character.scaleX)
        const scalesY = runs.map(character => character.scaleY)
        const rotationTerms = runs.map(character => character.matrix[1])
        const translationX = runs.map(character => character.matrix[4])
        const translationY = runs.map(character => character.matrix[5])
        assert.equal(runs.length, 4, `Shimmer must retain its four iterated characters: ${JSON.stringify(runs.map(character => character.text))}`)
        assert.ok(scalesX.every(value => value > 0 && value < 1) && scalesY.every(value => value > 0 && value < 1)
          && scalesX.every((value, index) => index === 0 || scalesX[index - 1] < value)
          && scalesY.every((value, index) => index === 0 || scalesY[index - 1] < value),
        `Shimmer scale tracks must stagger across characters: ${JSON.stringify({scalesX, scalesY})}`)
        assert.ok(rotationTerms.every(value => value < 0)
          && rotationTerms.every((value, index) => index === 0 || Math.abs(rotationTerms[index - 1]) > Math.abs(value)),
        `Shimmer rotation tracks must stagger across characters: ${JSON.stringify(rotationTerms)}`)
        assert.ok(translationX.every(value => value > 0)
          && translationX.every((value, index) => index === 0 || translationX[index - 1] > value)
          && translationY.every(value => value < 0)
          && translationY.every((value, index) => index === 0 || Math.abs(translationY[index - 1]) > Math.abs(value)),
        `Shimmer position tracks must stagger across characters: ${JSON.stringify({translationX, translationY})}`)
      }
      if (name === 'Shimmer' && targetMs === 2500) {
        assert.equal(state.runs.length, 4, `Shimmer must retain its four iterated characters after auto-reverse: ${JSON.stringify(state.runs.map(character => character.text))}`)
        const identity = [1, 0, 0, 1, 0, 0]
        assert.ok(state.runs.every(character => character.matrix.every((value, index) => Math.abs(value - identity[index]) < 0.01)),
          `Shimmer auto-reverse must restore every character to its original transform: ${JSON.stringify(state.runs.map(character => ({text: character.text, matrix: character.matrix})))}`)
      }
      if (name === 'Transparency') {
        close(target.opacity, 0.5, 0.001, 'Native image opacity filter must hold 50% opacity')
        assert.ok(!captureResult.warnings.some(warning => /image object animation.*not rendered|unsupported object-effect filter/.test(warning)), 'The native image opacity behavior must be recognized without an unsupported-filter fallback')
      }
      if (name === 'FlashBulb') {
        const progress = Math.min(1, targetMs / 2000)
        const fadeProgress = progress < 0.2 ? progress * 2.5 : progress <= 0.8 ? 0.5 : (1 - progress) * 2.5
        close(target.opacity, 1 - fadeProgress, 0.001, 'FlashBulb must return to full opacity at the end of its returning time filter')
        const {durationMs, scaleX, scaleY} = item.scaleTiming
        const cycleDurationMs = durationMs * 2
        const cycleElapsedMs = Math.min(targetMs, cycleDurationMs)
        const scaleProgress = cycleElapsedMs <= durationMs ? cycleElapsedMs / durationMs : 1 - (cycleElapsedMs - durationMs) / durationMs
        close(target.matrix[0], 1 + (scaleX - 1) * scaleProgress, 0.005, 'FlashBulb horizontal auto-reverse scale')
        close(target.matrix[3], 1 + (scaleY - 1) * scaleProgress, 0.005, 'FlashBulb vertical auto-reverse scale')
      }
      if (name === 'ChangeFontSize') {
        assert.ok(item.fontSizePt > 0, 'The native text must preserve a positive baseline font size')
        const baseFontSize = item.fontSizePt * frame.height / (slideHeight / 12700)
        close(run.fontSize, baseFontSize * 1.5 * Math.min(1, targetMs / 2000), 0.05, 'Native numeric to-only font-size emphasis')
      }
      if (name === 'ChangeFont') assert.ok(item.fontFamilyTo && run.fontFamily.includes(item.fontFamilyTo), `Native font-family p:set must use its exported destination: ${run.fontFamily}`)
      if (name === 'ChangeFontColor') {
        assert.ok(accent2Rgb, 'The native emphasis theme must preserve accent2')
        const actual = rgbChannels(run.color, 'Native font-color emphasis')
        const progress = Math.min(1, targetMs / 2000)
        accent2Rgb.forEach((value, index) => close(actual[index], Math.round(value * progress), 1, 'Native font-color channel'))
      }
      if (name === 'ChangeFontStyle') {
        assert.ok(item.fontStyleTo, 'The native ChangeFontStyle preset must export its fontStyle destination')
        assert.equal(run.fontStyle, item.fontStyleTo === 'italic' ? 'italic' : 'normal', 'ChangeFontStyle must apply the exported italic/normal style')
        if (item.fontWeightTo) assert.equal(Number(run.fontWeight), item.fontWeightTo === 'bold' ? 700 : 400, 'ChangeFontStyle must apply its exported font-weight p:set')
      }
      if (name === 'ChangeFillColor' && targetMs === 2500) {
        assert.ok(accent2Rgb, 'The native emphasis theme must preserve accent2')
        expectRgb(target.backgroundColor, accent2Rgb, 'ChangeFillColor accent2 fill')
      }
      if (name === 'ChangeLineColor' && targetMs === 2500) {
        assert.ok(accent2Rgb, 'The native emphasis theme must preserve accent2')
        expectRgb(target.strokeColor, accent2Rgb, 'ChangeLineColor accent2 stroke')
      }
      if (name === 'Spin') {
        const angle = Math.PI * Math.min(1, targetMs / 2000) * 2
        close(target.matrix[0], Math.cos(angle), 0.02, 'Native Spin rotation cosine')
        close(target.matrix[1], Math.sin(angle), 0.02, 'Native Spin rotation sine')
        close(target.matrix[2], -Math.sin(angle), 0.02, 'Native Spin rotation cross-axis')
        close(target.matrix[3], Math.cos(angle), 0.02, 'Native Spin rotation height')
      }
      if (name === 'GrowShrink') {
        const scale = 1 + 0.5 * Math.min(1, targetMs / 2000)
        close(target.scaleX, scale, 0.01, 'Native GrowShrink horizontal scale')
        close(target.scaleY, scale, 0.01, 'Native GrowShrink vertical scale')
      }
      if (name === 'GrowWithColor') {
        assert.ok(item.fontSizePt > 0 && item.staticTextColorRgb, 'GrowWithColor must retain its baseline font size and text color')
        assert.ok(accent2Rgb, 'GrowWithColor must preserve the theme accent2 destination')
        const baseFontSize = item.fontSizePt * frame.height / (slideHeight / 12700)
        assert.equal(state.runs.length, 4, 'GrowWithColor must preserve the four independently iterated characters')
        state.runs.forEach((character, index) => {
          const localElapsed = targetMs - item.childStyleDelayMs - index * item.iterationIntervalMs
          const progress = Math.min(1, Math.max(0, localElapsed / item.effectDurationMs))
          const expectedSize = baseFontSize * (1 + 0.5 * progress)
          close(character.fontSize, expectedSize, 0.05, `GrowWithColor character ${character.text} font-size track`)
          const expectedColor = item.staticTextColorRgb.map((channel, channelIndex) => Math.round(channel + (accent2Rgb[channelIndex] - channel) * progress))
          expectRgb(character.color, expectedColor, `GrowWithColor ${character.text} text-color track`)
        })
      }
      if (name === 'Lighten' || name === 'Desaturate') {
        assert.ok(!captureResult.warnings.some(warning => /color animation with an unsupported|generic color animation with missing|A color animation with missing/.test(warning)), `${name} native HSL color tracks must be recognized`)
        rgbChannels(run.color, `${name} native text color`)
        assert.equal(target.backgroundColor, 'rgba(0, 0, 0, 0)', `${name} must keep the exported no-fill text box transparent`)
        assert.equal(target.strokeColor, 'none', `${name} must not introduce an outline on the exported no-stroke text box`)
        if (name === 'Lighten') expectRgb(run.color, targetMs === 500 ? [0, 118, 236] : [77, 166, 255], 'Lighten from the saved blue text color')
        if (name === 'Desaturate') expectRgb(run.color, targetMs === 500 ? [18, 102, 186] : [72, 102, 132], 'Desaturate from the saved blue text color')
      }
      if (name === 'BoldFlash') assert.equal(Number(run.fontWeight), targetMs >= 1000 && targetMs < 1200 ? 700 : 400, 'BoldFlash must preserve its short discrete bold interval')
      if (name === 'BoldReveal') assert.equal(Number(run.fontWeight), 700, 'BoldReveal must hold bold text')
      if (name === 'BrushOnUnderline') assert.ok(run.textDecorationLine.includes('underline'), 'BrushOnUnderline must hold its exported underline')
      report.probes.push({name, targetMs, sampledMs: state.elapsedMs, result: 'PASS', evidence: `${name.toLowerCase()}-${targetMs}.png.json`})
      await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
      console.log(`${name} emphasis at ${targetMs}ms passed.`)
      continue
    }
    if (name === 'Swish') {
      assert.equal(item.rotationStartDegrees, -45, 'Swish must retain its -45-degree p:set starting rotation')
      if (targetMs === 500) {
        close(target.matrix[0], Math.cos(-Math.PI / 4), 0.02, 'Swish p:set rotation cosine')
        close(target.matrix[1], Math.sin(-Math.PI / 4), 0.02, 'Swish p:set rotation sine')
        close(target.matrix[2], -Math.sin(-Math.PI / 4), 0.02, 'Swish p:set rotation cross-axis')
        close(target.matrix[3], Math.cos(-Math.PI / 4), 0.02, 'Swish p:set rotation height')
      }
      assert.ok(!captureResult.warnings.some(warning => /p:set animation for style\.rotation/.test(warning)), 'Swish rotation p:set must not be reported as unrendered')
    }
    if (isWpsRuntime && metadata.Suite === 'CommonPresets' && name === 'Swish') {
      assert.ok(!captureResult.warnings.some(warning => /time-container timing curve or repeat without a finite positive duration is not applied|time-container timing curve or repeat is not applied because child timing exceeds its finite duration/.test(warning)),
        'WPS Swish must derive its parent curve from one text-unit timeline and preserve the character stagger')
    }
    if (name === 'Flip' && targetMs === 7000) {
      assert.equal(state.runs.length, 4, 'Flip must retain all four iterated text characters')
      for (const run of state.runs) {
        close(run.scaleX, 1, 0.015, `Flip ${run.text} must restore its original horizontal scale after auto-reverse`)
        close(run.scaleY, 1, 0.015, `Flip ${run.text} must restore its original vertical scale after auto-reverse`)
      }
    }
    if (['Iterated', 'Word'].includes(metadata.Suite)) {
      const runs = state.runs
      const expectedUnits = metadata.Suite === 'Word' ? ['ONE', 'TWO'] : ['A', 'B', 'C', 'D']
      assert.deepEqual(runs.map(run => run.text), expectedUnits, `${name} must render each native text unit exactly once`)
      if (metadata.Suite === 'Word' && targetMs === 2700) {
        const lineTops = runs.map(run => run.rect.y)
        assert.ok(Math.max(...lineTops) - Math.min(...lineTops) < 1,
          `${name} must preserve the WPS wrap="none" text line after word iteration: ${JSON.stringify({lineTops, runs: runs.map(run => ({text: run.text, rect: run.rect}))})}`)
      }
      const summary = effectSummary(runs)
      const nativeWipe = /^wipe\((right|left|up|down)\)$/.exec(item.effectFilter || '')
      const wipeCssDirection = nativeWipe ? ({right: 'right', left: 'left', up: 'bottom', down: 'top'})[nativeWipe[1]] : undefined
      if (name === 'Wipe' || name === 'WipeByWord') assert.ok(wipeCssDirection, `WPS must export a supported text-wipe direction, found: ${item.effectFilter}`)
      if (name === 'ColorByWord') {
        assert.ok(accent2Rgb, 'The native theme accent2 color must be available for the color animation reference')
        for (const [index, run] of runs.entries()) {
          const progress = Math.min(1, Math.max(0, (targetMs - index * item.iterationIntervalMs) / item.effectDurationMs))
          const expected = accent2Rgb.map(channel => Math.round(channel * progress))
          const actual = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(run.color || '')
          assert.ok(actual, `ColorByWord text unit ${run.text} must expose its rendered color: ${JSON.stringify(run)}`)
          for (let channel = 0; channel < 3; channel++) close(Number(actual[channel + 1]), expected[channel], 2, `ColorByWord ${run.text} ${['red', 'green', 'blue'][channel]} channel`)
          assert.equal(run.visibility, 'visible', `ColorByWord text unit ${run.text} must remain visible during its color emphasis`)
        }
        assert.ok(!captureResult.warnings.some(warning => /text iteration for color animation|unsupported text-range|color animation with an unsupported/.test(warning)), 'ColorByWord must be parsed without a skipped-animation warning')
        report.probes.push({name, targetMs, sampledMs: state.elapsedMs, iterationIntervalMs: item.iterationIntervalMs, effectDurationMs: item.effectDurationMs, letterStates: summary, result: 'PASS', evidence: `${name.toLowerCase()}-${targetMs}.png.json`})
        console.log(`${name} native color interpolation at ${targetMs}ms passed.`)
        await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
        continue
      }
      if (targetMs === 500) {
        for (const [index, run] of runs.entries()) {
          const progress = Math.min(1, Math.max(0, (targetMs - index * item.iterationIntervalMs) / item.effectDurationMs))
          assert.equal(run.visibility, progress > 0 ? 'visible' : 'hidden', `${name} letter ${run.text} must follow the native staggered start`)
          if (progress > 0 && name.startsWith('Fade')) close(run.opacity, progress, 0.015, `Fade text unit ${run.text} opacity`)
          if (progress > 0 && name === 'Zoom') {
            close(run.scaleX, progress, 0.015, `Zoom letter ${run.text} horizontal scale`)
            close(run.scaleY, progress, 0.015, `Zoom letter ${run.text} vertical scale`)
          }
        }
        if (!name.startsWith('Fade') && name !== 'Zoom') {
          const activeEffects = summary.slice(0, 3).map(run => run.effect)
          assert.ok(activeEffects.every(effect => effect !== 'none') && new Set(activeEffects).size === activeEffects.length,
            `${name} must render a distinct partial mask or clip for each active letter: ${JSON.stringify(summary)}`)
        }
        if (name === 'Wipe' || name === 'WipeByWord') {
          assert.ok(runs.every(run => run.maskImageInline.includes(`linear-gradient(to ${wipeCssDirection},`)),
            `Text wipe masks must preserve the WPS ${item.effectFilter} direction: ${JSON.stringify(runs.map(run => ({text: run.text, maskImageInline: run.maskImageInline, wipeProgress: run.wipeProgress})))}`)
        }
      } else {
        assert.ok(runs.every(run => run.visibility === 'visible'), `${name} must show every letter after the final staggered animation ends: ${JSON.stringify(summary)}`)
        if (name.startsWith('Fade')) assert.ok(runs.every(run => Math.abs(run.opacity - 1) < 0.015), 'Fade must finish fully opaque')
        if (name === 'Zoom') assert.ok(runs.every(run => Math.abs(run.scaleX - 1) < 0.015 && Math.abs(run.scaleY - 1) < 0.015), 'Zoom must finish at the original scale')
        if (name === 'Wipe' || name === 'WipeByWord') assert.ok(runs.every(run => run.wipeProgress === '100%' && run.maskImageInline.includes(`linear-gradient(to ${wipeCssDirection},`)), 'Text wipe must finish fully revealed while preserving its native direction')
      }
      assert.ok(!captureResult.warnings.some(warning => /unsupported text-range mask|generic .*formula.*skipped|unsupported time filter/.test(warning)), `${name} must not skip its native text effect`)
      report.probes.push({name, targetMs, sampledMs: state.elapsedMs, iterationIntervalMs: item.iterationIntervalMs, effectDurationMs: item.effectDurationMs, letterStates: summary, result: 'PASS', evidence: `${name.toLowerCase()}-${targetMs}.png.json`})
      console.log(`${name} iterated playback at ${targetMs}ms passed.`)
      await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
      continue
    }
    const progress = Math.min(1, state.elapsedMs / 2000)
    if (isWpsRuntime && metadata.Suite === 'CommonPresets' && name === 'Peek') {
      assert.ok(item.peekReference, 'WPS Peek must retain its native motion and wipe reference')
      if (targetMs < 2000) {
        close(target.matrix[5], item.peekReference.startOffsetRatio * target.rect.height * (1 - progress), 2, 'Peek shape-height motion track')
        assert.equal(target.wipeProgress, `${Math.round(progress * 100)}%`, 'Peek wipe progress must stay synchronized with its motion track')
        assert.ok(target.maskImageInline.includes('linear-gradient(to bottom,'), `Peek must reveal upward with the native wipe direction: ${JSON.stringify(target)}`)
      } else {
        close(target.matrix[5], 0, 0.25, 'Peek restored vertical position')
        assert.equal(target.wipeProgress, '100%', 'Peek must finish fully revealed')
      }
    }
    if (name === 'Bounce' && item.bounceReference) {
      for (const [index, run] of state.runs.entries()) {
        const unitElapsedMs = state.elapsedMs - index * item.bounceReference.iterationIntervalMs
        if (unitElapsedMs < 0) {
          assert.equal(run.visibility, 'hidden', `Bounce character ${run.text} must remain hidden before its iteration starts`)
          continue
        }
        const expected = bounceAt(item.bounceReference, unitElapsedMs)
        close(run.matrix[4] / frame.width, expected.x, 0.015, `Bounce ${run.text} horizontal time filter`)
        close(run.matrix[5] / frame.height, expected.y, 0.015, `Bounce ${run.text} vertical leg formula`)
        close(run.matrix[0], expected.scale[0], 0.045, `Bounce ${run.text} landing horizontal scale`)
        close(run.matrix[3], expected.scale[1], 0.045, `Bounce ${run.text} landing vertical scale`)
      }
    }
    if (['Blast', 'VerticalGrow'].includes(name)) {
      const actual = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(target.backgroundColor || '')
      assert.ok(actual, `${name} must reveal a solid shape fill: ${JSON.stringify({backgroundColor: target.backgroundColor, warnings: captureResult.warnings})}`)
      assert.ok(!captureResult.warnings.some(warning => /p:set animation for fill\.type|fill\.on p:set animation without/.test(warning)), `${name} solid-fill setup must be parsed without a skipped-animation warning`)
      if (name === 'Blast') {
        const fillProgress = Math.min(1, Math.max(0, (state.elapsedMs - item.fillColorTiming.delayMs) / item.fillColorTiming.durationMs))
        const expected = accent2Rgb.map(channel => Math.round(channel * fillProgress))
        for (let channel = 0; channel < 3; channel++) close(Number(actual[channel + 1]), expected[channel], 2, `Blast fill ${['red', 'green', 'blue'][channel]} channel`)
      }
      if (name === 'VerticalGrow' && targetMs === 4000) {
        for (let channel = 0; channel < 3; channel++) close(Number(actual[channel + 1]), 0, 1, `VerticalGrow reverse fill ${['red', 'green', 'blue'][channel]} channel`)
      }
    }
    if (targetMs < 2000) {
      const randomEffectHasIntermediateVisual = name === 'RandomEffects' && item.randomEffectPlaceholder && (
        target.opacity < 0.999 || target.maskImage !== 'none' || target.clipPath !== 'none'
          || Math.abs(target.matrix[0] - 1) > 0.02 || Math.abs(target.matrix[1]) > 0.02
          || Math.abs(target.matrix[2]) > 0.02 || Math.abs(target.matrix[3] - 1) > 0.02
          || Math.abs(target.matrix[4]) > 1 || Math.abs(target.matrix[5]) > 1
      )
      if (!randomEffectHasIntermediateVisual) assert.equal(target.visibility, 'visible', `${name} must be visible during playback`)
      if (name === 'RandomEffects') {
        const hasIntermediateVisual = (target.opacity > 0.001 && target.opacity < 0.999)
          || target.maskImage !== 'none'
          || target.clipPath !== 'none'
          || Math.abs(target.matrix[0] - 1) > 0.02
          || Math.abs(target.matrix[1]) > 0.02
          || Math.abs(target.matrix[2]) > 0.02
          || Math.abs(target.matrix[3] - 1) > 0.02
          || Math.abs(target.matrix[4]) > 1
          || Math.abs(target.matrix[5]) > 1
        assert.ok(hasIntermediateVisual, `RandomEffects must render a visible intermediate effect: ${JSON.stringify(target)}`)
        if (item.randomEffectPlaceholder) assert.equal(target.visibility, 'visible', 'RandomEffects must be visible while its entrance fallback is playing')
        if (item.randomEffectPlaceholder) assert.ok(captureResult.warnings.some(warning => /Random Effects preset omitted its selected effect/.test(warning)), 'RandomEffects must disclose the chosen fallback for its omitted effect')
        else assert.ok(!captureResult.warnings.some(warning => /Random Effects preset omitted its selected effect/.test(warning)), 'A concrete native RandomEffects effect must not be replaced by the fallback')
        if (!isWpsRuntime) assert.ok(!captureResult.warnings.some(warning => /WPS Random Effects/.test(warning)), 'PowerPoint playback warnings must not name WPS as the source')
        assert.ok(!captureResult.warnings.some(warning => /generic object animation with an unsupported target/.test(warning)), 'RandomEffects must not skip the placeholder as a generic unsupported animation')
        randomEffectMidStates.set(name, JSON.stringify([target.opacity, target.maskImage, target.clipPath, target.matrix]))
      }
      if (name === 'Fly') close(target.matrix[5] / frame.height, (1 - item.y / slideHeight) * (1 - progress), 0.025, 'Fly starts from the bottom edge')
      if (name === 'Spin') {
        close(target.matrix[0], Math.cos(2 * Math.PI * progress), 0.08, 'Spin cosine')
        close(target.matrix[1], Math.sin(2 * Math.PI * progress), 0.08, 'Spin sine')
      }
      if (name === 'GrowShrink') close(target.matrix[0], 1 + progress * 0.5, 0.03, 'GrowShrink declared scale')
      if (name === 'GrowAndTurn') {
        assert.deepEqual(item.rotationKeyframes, [90, 0], 'GrowAndTurn reference values must remain in degrees')
        close(Math.atan2(target.matrix[1], target.matrix[0]), Math.PI / 2 * (1 - progress), 0.08, 'GrowAndTurn rotation angle')
      }
      if (name === 'Swivel' || name === 'FadedSwivel') {
        close(target.matrix[0], Math.sin(2.5 * Math.PI * progress), 0.08, `${name} signed width`)
        close(target.matrix[3], 1, 0.01, `${name} unchanged height`)
        if (name === 'FadedSwivel') close(target.opacity, progress, 0.04, 'FadedSwivel combines opacity and signed scale')
      }
      if (name === 'FadeExit') close(target.opacity, 1 - progress, 0.04, 'FadeExit opacity')
      if (name === 'ZoomExit') {
        close(target.matrix[0], 1 + 3 * progress, 0.05, 'ZoomExit width')
        close(target.matrix[3], 1 + 3 * progress, 0.05, 'ZoomExit height')
      }
      if (name === 'FlyExit') close(target.matrix[5] / frame.height, (1 - item.y / slideHeight) * progress, 0.025, 'FlyExit destination')
    } else if (name.endsWith('Exit')) {
      const targets = item.iterated ? state.runs : [target]
      assert.ok(targets.every(element => element.visibility === 'hidden'), `${name} must hide after its exit`)
      if (name === 'ZoomExit') assert.ok(targets.every(element => Math.abs(element.matrix[0] - 4) < 0.01 && Math.abs(element.matrix[3] - 4) < 0.01), 'ZoomExit must retain its declared terminal scale')
      if (name === 'FlyExit' && !item.iterated) close(target.rect.y, frame.y + frame.height, 2, 'FlyExit ends immediately below the slide')
    } else {
      const targets = item.iterated ? state.runs : [target]
      if (metadata.Suite === 'CommonPresets' && name === 'Swish' && targetMs === 2700) {
        if (isWpsRuntime) {
          assert.deepEqual(targets.map(element => [element.text, element.visibility]), [['S', 'visible'], ['w', 'visible'], ['i', 'visible'], ['s', 'hidden'], ['h', 'hidden']],
            'WPS Swish must preserve the observed three-of-five text-unit state at 2700ms')
        } else {
          assert.ok(targets.some(element => element.visibility === 'visible') && targets.some(element => element.visibility === 'hidden'), 'Swish must preserve staggered text playback at 2700ms')
        }
      } else assert.ok(targets.every(element => element.visibility === 'visible'), `${name} must remain visible after its entrance`)
      if (name === 'RandomEffects') {
        if (item.randomEffectPlaceholder) assert.ok(captureResult.warnings.some(warning => /Random Effects preset omitted its selected effect/.test(warning)), 'RandomEffects must disclose the chosen fallback for its omitted effect')
        else assert.ok(!captureResult.warnings.some(warning => /Random Effects preset omitted its selected effect/.test(warning)), 'A concrete native RandomEffects effect must not be replaced by the fallback')
        assert.ok(!captureResult.warnings.some(warning => /generic object animation with an unsupported target/.test(warning)), 'RandomEffects must not skip the placeholder as a generic unsupported animation')
        assert.notEqual(JSON.stringify([target.opacity, target.maskImage, target.clipPath, target.matrix]), randomEffectMidStates.get(name), 'RandomEffects must finish its intermediate effect and settle into its final visible state')
      }
      if (name === 'Bounce') assert.ok(targets.every(element => Math.abs(element.matrix[0] - 1) < 0.01 && Math.abs(element.matrix[3] - 1) < 0.01 && Math.abs(element.matrix[4]) < 0.1 && Math.abs(element.matrix[5]) < 0.1), 'Bounce must restore every letter to its saved position and scale')
      if (isWpsRuntime && metadata.Suite === 'CommonPresets' && name === 'Swish' && targetMs === 6000) {
        for (const run of targets) {
          close(run.matrix[0], 1, 0.02, `Swish ${run.text} restored horizontal transform`)
          close(run.matrix[1], 0, 0.02, `Swish ${run.text} restored vertical transform`)
          close(run.matrix[2], 0, 0.02, `Swish ${run.text} restored cross-axis transform`)
          close(run.matrix[3], 1, 0.02, `Swish ${run.text} restored vertical scale`)
        }
      }
      if (metadata.Suite === 'CommonPresets' && name === 'VerticalGrow' && targetMs === 4000) {
        close(target.matrix[0], 1, 0.02, 'VerticalGrow restores its original width')
        close(target.matrix[3], 1, 0.02, 'VerticalGrow restores its original height')
      }
    }
    assert.ok(!captureResult.warnings.some(warning => /generic .*formula.*skipped|unsupported time filter|generic object animation with an unsupported target/.test(warning)), `${name} must not skip its native property tracks`)
    report.probes.push({name, iterated: item.iterated, targetMs, sampledMs: state.elapsedMs, wallElapsedMs: state.wallElapsedMs, ...(motionPathCheck ? {motionPathCheck} : {}), result: 'PASS', evidence: `${name.toLowerCase()}-${targetMs}.png.json`})
    console.log(`${name} at ${Math.round(state.elapsedMs)}ms passed.`)
    await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
  }
    report.result = 'PASS'
} catch (error) {
  report.result = 'FAIL'
  report.error = error instanceof Error ? error.message : String(error)
  throw error
} finally {
  await writeFile(join(outputRoot, 'playback-verification.json'), JSON.stringify(report, null, 2))
}
console.log(`Native Office/WPS ${metadata.Suite.toLowerCase()}-suite browser playback verification passed.`)
