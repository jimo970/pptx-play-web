import JSZip from 'jszip'
import {inverseTimingProgress, type TimingCurve} from './animation-timing'
import {isSupportedChartNumberFormat} from './chart-number-format'
import {hslToRgb} from './color'

export interface PptxRun {
    text: string
    fontFamily?: string
    fontFamilyEastAsia?: string
    fontFamilyComplexScript?: string
    fontSizePt?: number
    letterSpacingPt?: number
    baselineOffsetEm?: number
    color?: string
    outlineColor?: string
    outlineWidth?: number
    bold?: boolean
    italic?: boolean
    underline?: boolean
    underlineStyle?: 'double' | 'dotted' | 'dashed' | 'wavy'
    lineThrough?: boolean
    lineThroughStyle?: 'solid' | 'double'
    hyperlink?: PptxHyperlink
    shadow?: PptxShadow
}

export type PptxHyperlink =
    | {kind: 'external'; url: string; tooltip?: string}
    | {kind: 'slide'; targetPath: string; tooltip?: string}
    | {kind: 'showJump'; jump: 'firstslide' | 'lastslide' | 'nextslide' | 'previousslide' | 'number'; slideNumber?: number; tooltip?: string}

export interface PptxShadow {
    color: string
    blur: number
    distance: number
    angle: number
}

export interface PptxParagraph {
    align: 'left' | 'center' | 'right' | 'justify'
    runs: PptxRun[]
    level?: number
    marginLeft?: number
    marginRight?: number
    indent?: number
    spaceBefore?: {unit: 'emu' | 'ratio'; value: number}
    spaceAfter?: {unit: 'emu' | 'ratio'; value: number}
    lineSpacing?: {unit: 'emu' | 'ratio'; value: number}
    bullet?: {kind: 'char'; char: string} | {kind: 'number'; type: string; startAt?: number}
}

export interface PptxTableCell {
    paragraphs: PptxParagraph[]
    textColumnCount?: number
    textColumnSpacing?: number
    textFontScale?: number
    textLineSpacingReduction?: number
    textAutoFitDynamic?: boolean
    fill: string
    margins: {left: number; right: number; top: number; bottom: number}
    verticalAlign: 'top' | 'middle' | 'bottom'
    colSpan: number
    rowSpan: number
    hidden: boolean
    borders: Partial<Record<'top' | 'right' | 'bottom' | 'left', {color: string; width: number}>>
}

export interface PptxTable {
    columns: number[]
    rows: {height: number; cells: PptxTableCell[]}[]
}

export type PptxChartAxisPosition = 'b' | 'l' | 'r' | 't'
export type PptxChartAxisCrosses = 'autoZero' | 'max' | 'min'
export type PptxChartTickMark = 'cross' | 'in' | 'none' | 'out'
export type PptxChartTickLabelPosition = 'high' | 'low' | 'nextTo' | 'none'
export type PptxChartDataLabelPosition = 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't'

export interface PptxChartSeries {
    legendIndex: number
    name: string
    values: Array<number | null>
    color: string
    pointColors: Record<number, string>
    showValueLabels: boolean
    valueLabelOverrides: Record<number, boolean>
    showPercentLabels: boolean
    percentLabelOverrides: Record<number, boolean>
    showCategoryNameLabels: boolean
    categoryNameLabelOverrides: Record<number, boolean>
    showSeriesNameLabels: boolean
    seriesNameLabelOverrides: Record<number, boolean>
    labelSeparator?: string
    labelSeparatorOverrides: Record<number, string>
    numberFormat?: string
    valueNumberFormatOverrides: Record<number, string>
    dataLabelPosition?: PptxChartDataLabelPosition
    dataLabelPositionOverrides: Record<number, PptxChartDataLabelPosition>
}

export interface PptxChart {
    type: 'bar' | 'line' | 'pie' | 'doughnut'
    direction?: 'horizontal' | 'vertical'
    holeSize?: number
    firstSliceAngle?: number
    title?: string
    categories: string[]
    palette: string[]
    series: PptxChartSeries[]
    labelSeparator?: string
    numberFormat?: string
    valueAxis?: {
        min: number
        max: number
        majorUnit?: number
        numberFormat?: string
        reverse?: boolean
        deleted?: boolean
        majorGridlines?: boolean
        position?: PptxChartAxisPosition
        crosses?: PptxChartAxisCrosses
        crossesAt?: number
        majorTickMark?: PptxChartTickMark
        tickLabelPosition?: PptxChartTickLabelPosition
    }
    categoryAxisPosition?: PptxChartAxisPosition
    categoryAxisCrosses?: PptxChartAxisCrosses
    categoryAxisCrossesAt?: number
    categoryAxisDeleted?: boolean
    categoryAxisMajorTickMark?: PptxChartTickMark
    categoryAxisTickLabelPosition?: PptxChartTickLabelPosition
    legend?: {
        position: 'bottom' | 'top' | 'left' | 'right' | 'topRight'
        hiddenEntries?: number[]
        overlay?: boolean
        manualLayout?: {
            x?: {value: number; mode: 'edge' | 'factor'}
            y?: {value: number; mode: 'edge' | 'factor'}
            w?: {value: number; mode: 'edge' | 'factor'}
            h?: {value: number; mode: 'edge' | 'factor'}
        }
    }
    dataLabelPosition?: PptxChartDataLabelPosition
}

export interface PptxCustomPath {
    d: string
    fill: boolean
    stroke: boolean
}

type PptxTextOrientation = 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'
type PptxLineEnd = {type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow'; width: 'sm' | 'med' | 'lg'; length: 'sm' | 'med' | 'lg'}

export interface PptxElement {
    id: string
    name: string
    kind: 'shape' | 'picture' | 'video' | 'audio' | 'line' | 'placeholder' | 'table' | 'chart'
    geometry: string
    customPaths?: PptxCustomPath[]
    customTextRect?: {left: number; top: number; right: number; bottom: number}
    x: number
    y: number
    width: number
    height: number
    rotation: number
    flipH: boolean
    flipV: boolean
    renderMatrix?: [number, number, number, number, number, number]
    motionParent?: {matrix: [number, number, number, number, number, number]; x: number; y: number}
    fill: string
    stroke: string
    strokeWidth: number
    strokeDash?: string | number[]
    strokeCap?: 'butt' | 'round' | 'square'
    lineHeadEnd?: PptxLineEnd
    lineTailEnd?: PptxLineEnd
    shadow?: PptxShadow
    imageUrl?: string
    mediaUrl?: string
    mediaType?: 'video' | 'audio'
    mediaMime?: string
    mediaFormat?: string
    mediaTrim?: {startMs: number; endMs: number}
    mediaFade?: {inMs: number; outMs: number}
    imageCrop?: {left: number; top: number; right: number; bottom: number}
    table?: PptxTable
    chart?: PptxChart
    paragraphs: PptxParagraph[]
    margins: {left: number; right: number; top: number; bottom: number}
    verticalAlign: 'top' | 'middle' | 'bottom'
    textWrap?: 'square' | 'none'
    textOrientation?: PptxTextOrientation
    textColumnCount?: number
    textColumnSpacing?: number
    textFontScale?: number
    textLineSpacingReduction?: number
    textAutoFit?: 'normal' | 'shape'
    textAutoFitDynamic?: boolean
    placeholder?: {type: string; index?: string}
}

export interface PptxSlide {
    id: string
    name: string
    elements: PptxElement[]
    background: string
    backgroundImage?: {
        url: string
        crop: {left: number; top: number; right: number; bottom: number}
        tile?: {scaleX: number; scaleY: number; align: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br'; flip: 'none' | 'x' | 'y' | 'xy'; offsetX: number; offsetY: number}
    }
    speakerNotes?: string
    automaticAnimations?: PptxAnimation[]
    animationSteps?: PptxAnimation[][]
    animationSequence?: {
        nextAction: 'none' | 'seek'
        advancesOnNext?: boolean
        rewindsOnPrevious?: boolean
    }
    triggeredAnimations?: {
        trigger: {type: 'shape'; id: string; event: 'onClick' | 'onDblClick' | 'onMouseOver' | 'onMouseOut'} | {type: 'slide'; event: 'onNext' | 'onPrev'}
        steps: PptxAnimation[][]
    }[]
    advanceOnClick?: boolean
    advanceAfterMs?: number
    transition?: {
        effect:
            | 'fade'
            | 'push'
            | 'wipe'
            | 'zoom'
            | 'cover'
            | 'pull'
            | 'split'
            | 'circle'
            | 'diamond'
            | 'plus'
            | 'wedge'
            | 'newsflash'
            | 'flash'
            | 'doors'
            | 'window'
            | 'prism'
            | 'pan'
            | 'vortex'
            | 'ferris'
            | 'gallery'
            | 'conveyor'
            | 'ripple'
            | 'glitter'
            | 'shred'
            | 'flythrough'
            | 'warp'
            | 'flip'
            | 'switch'
            | 'reveal'
            | 'morph'
            | 'preset'
            | 'random'
            | 'cut'
            | 'honeycomb'
            | 'blinds'
            | 'checker'
            | 'comb'
            | 'dissolve'
            | 'wheel'
            | 'wheelReverse'
            | 'randomBar'
            | 'strips'
        morphOption?: 'byObject' | 'byWord' | 'byChar'
        presetName?: 'fallOver' | 'drape' | 'curtains' | 'wind' | 'prestige' | 'fracture' | 'crush' | 'peelOff' | 'pageCurlDouble' | 'pageCurlSingle' | 'airplane' | 'origami'
        invertX?: boolean
        invertY?: boolean
        pattern?: 'diamond' | 'hexagon'
        shredPattern?: 'strip' | 'rectangle'
        hasBounce?: boolean
        direction: string
        orientation?: 'horz' | 'vert'
        barOrientation?: 'horizontal' | 'vertical'
        spokes?: number
        isInverted?: boolean
        isContent?: boolean
        throughBlack?: boolean
        durationMs: number
    }
}

export interface PptxTransitionPlayback {
    slideId: string
    startedAtEpoch: number
    transition: NonNullable<PptxSlide['transition']>
}

export interface PptxAnimation {
    targetId: string
    paragraphRange?: {start: number; end: number}
    characterRange?: {start: number; end: number}
    iteration?: {intervalMs: number; ranges: Array<{start: number; end: number}>; backwards?: boolean}
    effect: 'appear' | 'fade' | 'wipe' | 'blinds' | 'checker' | 'randomBars' | 'strips' | 'barn' | 'dissolve' | 'shape' | 'wheel' | 'slide' | 'scale' | 'motion' | 'timingOnly' | 'rotation' | 'color' | 'fontSize' | 'fontWeight' | 'opacity' | 'media'
    direction: 'in' | 'out'
    durationMs: number
    delayMs: number
    fillMode?: 'hold' | 'freeze' | 'remove'
    sequenceKey?: string
    sequenceOrder?: number
    repeatCount?: number
    repeatDurationMs?: number
    autoReverse?: boolean
    wipeDirection?: 'right' | 'left' | 'up' | 'down'
    blindsOrientation?: 'horizontal' | 'vertical'
    checkerOrientation?: 'horz' | 'vert'
    randomBarOrientation?: 'horizontal' | 'vertical'
    stripsDirection?: 'ld' | 'lu' | 'rd' | 'ru'
    barnOrientation?: 'horizontal' | 'vertical'
    barnMotion?: 'in' | 'out'
    shapeFilter?: 'circle' | 'diamond' | 'box' | 'plus'
    shapeDirection?: 'in' | 'out'
    slideFrom?: 'left' | 'right' | 'top' | 'bottom'
    wheelSpokes?: 1 | 2 | 3 | 4 | 8
    scaleFrom?: [number, number]
    scaleTo?: [number, number]
    rotationFrom?: number
    rotationTo?: number
    rotationRelative?: boolean
    colorProperty?: 'fillcolor' | 'style.color' | 'stroke.color' | 'shadow.color'
    opacityProperty?: 'style.opacity' | 'fill.opacity' | 'stroke.opacity' | 'shadow.opacity'
    colorFrom?: [number, number, number, number]
    colorTo?: [number, number, number, number]
    colorSpace?: 'rgb' | 'hsl'
    colorDirection?: 'cw' | 'ccw'
    fontSizeFrom?: number
    fontSizeTo?: number
    fontSizeKeyframes?: Array<{offset: number; value: number}>
    fontSizeKeyframeMode?: 'lin' | 'discrete'
    fontWeightFrom?: number
    fontWeightTo?: number
    fontWeightKeyframes?: Array<{offset: number; value: number}>
    fontWeightKeyframeMode?: 'lin' | 'discrete'
    fontWeightFormulaSamples?: Array<{offset: number; value: number}>
    opacityFrom?: number
    opacityTo?: number
    opacityKeyframes?: Array<{offset: number; value: number}>
    opacityKeyframeMode?: 'lin' | 'discrete'
    mediaCommand?: 'play' | 'pause' | 'togglePause' | 'stop'
    mediaActionKey?: string
    mediaDurationMs?: number
    mediaSlideCount?: number
    mediaWaitForEnd?: boolean
    mediaWaitForEndKeys?: string[]
    mediaWaitDelayMs?: number
    mediaStartSeconds?: number
    mediaVolume?: number
    mediaMuted?: boolean
    motionPath?: Array<{x: number; y: number; distance: number}>
    motionPathLength?: number
    motionPathSamples?: Array<{progress: number; x: number; y: number}>
    acceleration?: number
    deceleration?: number
    playbackSpeed?: number
    reversePlayback?: boolean
    timingWarp?: TimingCurve & {
        groupOffsetMs: number
        groupDurationMs: number
        groupSpeed: number
        groupReversePlayback: boolean
        groupAutoReverse: boolean
        groupPlaybackDurationMs: number
        childOffsetMs: number
        groupEndDelayMs: number
    }
}

export interface PptxTriggerPlayback {
    sequence: number
    step: number
    startedAt: number
}

interface PptxRelationship {
    target: string
    type: string
    external?: boolean
}

export interface PptxDocument {
    name: string
    width: number
    height: number
    slides: PptxSlide[]
    warnings: string[]
    objectUrls: string[]
}

type PptxTextStyle = Omit<PptxRun, 'text'>
type ParagraphValues = Partial<Omit<PptxParagraph, 'runs'>>
interface ParagraphDefaults {
    values: ParagraphValues
    text: PptxTextStyle
}
type ParagraphDefaultsByLevel = Record<number, ParagraphDefaults>
interface MasterTextStyle {
    style: PptxTextStyle
    paragraphs: ParagraphDefaultsByLevel
}

interface ParsedPlaceholder {
    node: Element
    type: string
    index?: string
    element: PptxElement
    paragraphStyles: ParagraphDefaultsByLevel
}

const MAX_FILE_BYTES = 150 * 1024 * 1024
const MAX_ARCHIVE_BYTES = 250 * 1024 * 1024
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PRESENTATION_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const OLE_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/presentationml/2006/ole'
const POWERPOINT_2010_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
const POWERPOINT_2012_NS = 'http://schemas.microsoft.com/office/powerpoint/2012/main'
const POWERPOINT_2015_NS = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
const SUPPORTED_MC_REQUIRES_NAMESPACES = new Set([PRESENTATION_NS, DRAWING_NS, POWERPOINT_2010_NS, POWERPOINT_2012_NS, POWERPOINT_2015_NS])
const SUPPORTED_TRANSITIONS = new Set([
    'fade',
    'push',
    'wipe',
    'zoom',
    'cover',
    'pull',
    'split',
    'circle',
    'diamond',
    'plus',
    'wedge',
    'newsflash',
    'flash',
    'doors',
    'window',
    'prism',
    'pan',
    'vortex',
    'ferris',
    'gallery',
    'conveyor',
    'ripple',
    'glitter',
    'shred',
    'flythrough',
    'warp',
    'flip',
    'switch',
    'reveal',
    'morph',
    'prstTrans',
    'random',
    'cut',
    'honeycomb',
    'blinds',
    'checker',
    'comb',
    'dissolve',
    'wheel',
    'wheelReverse',
    'randomBar',
    'strips'
])
// p15:prstTrans preset names and default-false invX/invY attributes follow MS-PPTX CT_PresetTransition.
const SUPPORTED_PRESET_TRANSITIONS = new Set(['fallOver', 'drape', 'curtains', 'wind', 'prestige', 'fracture', 'crush', 'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami'])
const DEFAULT_THEME: Record<string, string> = {
    dk1: '#000000',
    lt1: '#ffffff',
    dk2: '#1f497d',
    lt2: '#eeece1',
    accent1: '#4472c4',
    accent2: '#ed7d31',
    accent3: '#a5a5a5',
    accent4: '#ffc000',
    accent5: '#5b9bd5',
    accent6: '#70ad47',
    hlink: '#0563c1',
    folHlink: '#954f72'
}

function children(node: Element | null | undefined, name?: string): Element[] {
    if (!node) return []
    return Array.from(node.children).filter(child => !name || child.localName === name)
}

function child(node: Element | null | undefined, name: string): Element | undefined {
    return children(node, name)[0]
}

function descendants(node: Element | null | undefined, name: string): Element[] {
    if (!node) return []
    return Array.from(node.getElementsByTagName('*')).filter(item => item.localName === name)
}

function firstDescendant(node: Element | null | undefined, name: string): Element | undefined {
    return descendants(node, name)[0]
}

function parseXml(source: string, label: string): XMLDocument {
    try {
        const xml = new DOMParser().parseFromString(source, 'application/xml')
        if (xml.getElementsByTagName('parsererror').length) throw new Error()
        return xml
    } catch {
        throw new Error(`${label} contains invalid XML.`)
    }
}

function numberAttr(node: Element | undefined, key: string, fallback = 0): number {
    const raw = node?.getAttribute(key)
    if (!raw) return fallback
    const value = Number(raw)
    return Number.isFinite(value) ? value : fallback
}

function fixedPercentage(value: string | null): number | undefined {
    if (value === null) return undefined
    const raw = value.trim()
    const percent = /^(?:\d+\.?\d*|\.\d+)%$/.test(raw) ? Number(raw.slice(0, -1)) / 100 : /^\d+$/.test(raw) ? Number(raw) / 100_000 : Number.NaN
    return Number.isFinite(percent) && percent >= 0 && percent <= 1 ? percent : undefined
}

function namespacedAttr(node: Element | undefined, name: string): string | null {
    if (!node) return null
    return node.getAttributeNS(OFFICE_REL_NS, name) || node.getAttribute(`r:${name}`) || Array.from(node.attributes).find(attribute => attribute.namespaceURI === OFFICE_REL_NS && attribute.localName === name)?.value || null
}

function resolveTarget(sourcePath: string, target: string): string {
    const path = target.replaceAll('\\', '/').replace(/^\//, '')
    const parts = path.startsWith('ppt/') ? [] : sourcePath.split('/').slice(0, -1)
    for (const part of path.split('/')) {
        if (!part || part === '.') continue
        if (part === '..') parts.pop()
        else parts.push(part)
    }
    return parts.join('/')
}

function relationshipPath(sourcePath: string): string {
    const slash = sourcePath.lastIndexOf('/')
    const folder = slash < 0 ? '' : sourcePath.slice(0, slash + 1)
    const filename = slash < 0 ? sourcePath : sourcePath.slice(slash + 1)
    return `${folder}_rels/${filename}.rels`
}

async function loadRelationships(zip: JSZip, path: string): Promise<Map<string, PptxRelationship>> {
    const entry = zip.file(path)
    if (!entry) return new Map()
    const xml = parseXml(await entry.async('string'), path)
    return new Map(
        Array.from(xml.getElementsByTagName('*'))
            .filter(node => node.localName === 'Relationship')
            .map(node => {
                const external = node.getAttribute('TargetMode') === 'External'
                return [
                    node.getAttribute('Id') || '',
                    {
                        target: external ? node.getAttribute('Target') || '' : resolveTarget(path.replace(/_rels\/[^/]+\.rels$/, path.split('/').at(-1) || ''), node.getAttribute('Target') || ''),
                        type: node.getAttribute('Type') || '',
                        external
                    }
                ]
            })
    )
}

function validateArchiveSize(zip: JSZip): void {
    let total = 0
    for (const entry of Object.values(zip.files)) {
        if (entry.dir) continue
        // ponytail: JSZip keeps expanded size private; replace this guard if it adds public metadata.
        const size = (entry as JSZip.JSZipObject & {_data?: {uncompressedSize?: number}})._data?.uncompressedSize
        if (!Number.isFinite(size)) throw new Error('This presentation has incomplete ZIP size metadata and cannot be read safely.')
        total += size!
        if (total > MAX_ARCHIVE_BYTES) throw new Error('This presentation expands beyond the 250 MB browser safety limit.')
    }
}

function rgbToCss(hex: string, alpha = 1): string {
    const cleaned = hex.replace('#', '')
    if (!/^[0-9a-f]{6}$/i.test(cleaned)) return '#000000'
    const r = parseInt(cleaned.slice(0, 2), 16)
    const g = parseInt(cleaned.slice(2, 4), 16)
    const b = parseInt(cleaned.slice(4, 6), 16)
    return alpha >= 0.999 ? `#${cleaned}` : `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(alpha, 1))})`
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
    r = Math.max(0, Math.min(r / 255, 1))
    g = Math.max(0, Math.min(g / 255, 1))
    b = Math.max(0, Math.min(b / 255, 1))
    const max = Math.max(r, g, b),
        min = Math.min(r, g, b),
        delta = max - min
    const lightness = (max + min) / 2
    if (!delta) return [0, 0, lightness]
    const saturation = delta / (1 - Math.abs(2 * lightness - 1))
    let hue = max === r ? ((g - b) / delta) % 6 : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4
    hue = (hue / 6 + 1) % 1
    return [hue, saturation, lightness]
}

function readColor(node: Element | undefined, theme: Record<string, string>): string | undefined {
    if (!node) return undefined
    let hex = ''
    if (node.localName === 'srgbClr') hex = node.getAttribute('val') || ''
    else if (node.localName === 'sysClr') hex = node.getAttribute('lastClr') || '000000'
    else if (node.localName === 'schemeClr') hex = theme[node.getAttribute('val') || ''] || '#000000'
    else if (node.localName === 'hslClr') {
        const [r, g, b] = hslToRgb(numberAttr(node, 'hue') / 21_600_000, numberAttr(node, 'sat', 100_000) / 100_000, numberAttr(node, 'lum', 50_000) / 100_000)
        hex = [r, g, b].map(channel => Math.round(channel).toString(16).padStart(2, '0')).join('')
    } else if (node.localName === 'prstClr') {
        const preset: Record<string, string> = {black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080', grey: '#808080'}
        hex = preset[node.getAttribute('val') || ''] || '#000000'
    }
    if (!hex) return undefined
    if (!hex.startsWith('#')) hex = `#${hex}`
    let r = parseInt(hex.slice(1, 3), 16)
    let g = parseInt(hex.slice(3, 5), 16)
    let b = parseInt(hex.slice(5, 7), 16)
    let alpha = 1
    for (const mod of children(node)) {
        const value = numberAttr(mod, 'val', 100_000) / 100_000
        if (mod.localName === 'alpha') alpha = value
        else if (mod.localName === 'tint') {
            r += (255 - r) * value
            g += (255 - g) * value
            b += (255 - b) * value
        } else if (mod.localName === 'shade') {
            r *= value
            g *= value
            b *= value
        } else if (mod.localName === 'lumMod') {
            // DrawingML luminance transforms operate on HSL lightness, not RGB channels.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminancemodulation
            const [hue, saturation, lightness] = rgbToHsl(r, g, b)
            ;[r, g, b] = hslToRgb(hue, saturation, Math.max(0, Math.min(lightness * value, 1)))
        } else if (mod.localName === 'lumOff') {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminanceoffset
            const [hue, saturation, lightness] = rgbToHsl(r, g, b)
            ;[r, g, b] = hslToRgb(hue, saturation, Math.max(0, Math.min(lightness + value, 1)))
        }
    }
    return rgbToCss(
        [r, g, b]
            .map(channel =>
                Math.round(Math.max(0, Math.min(255, channel)))
                    .toString(16)
                    .padStart(2, '0')
            )
            .join(''),
        alpha
    )
}

function animationColorChannels(node: Element | undefined, theme: Record<string, string>, colorSpace: 'rgb' | 'hsl'): [number, number, number, number] | undefined {
    const css = readColor(node, theme)
    if (!css) return undefined
    const hex = /^#([0-9a-f]{6})$/i.exec(css)
    const rgba = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*(?:,\s*([\d.]+)\s*)?\)$/i.exec(css)
    const channels = hex
        ? ([parseInt(hex[1]!.slice(0, 2), 16), parseInt(hex[1]!.slice(2, 4), 16), parseInt(hex[1]!.slice(4, 6), 16), 1] as [number, number, number, number])
        : rgba
          ? ([Number(rgba[1]), Number(rgba[2]), Number(rgba[3]), rgba[4] === undefined ? 1 : Number(rgba[4])] as [number, number, number, number])
          : undefined
    if (!channels || colorSpace === 'rgb') return channels
    const [hue, saturation, lightness] = rgbToHsl(channels[0], channels[1], channels[2])
    return [hue, saturation, lightness, channels[3]]
}

function outerShadow(parent: Element | undefined, theme: Record<string, string>): PptxShadow | undefined {
    const shadow = firstDescendant(parent, 'outerShdw')
    if (!shadow) return undefined
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow
    return {
        color: readColor(children(shadow)[0], theme) || 'rgba(0, 0, 0, .35)',
        blur: numberAttr(shadow, 'blurRad'),
        distance: numberAttr(shadow, 'dist'),
        angle: numberAttr(shadow, 'dir') / 60_000
    }
}

function fillColor(parent: Element | undefined, theme: Record<string, string>, warnings: Set<string>): string {
    const fill = children(parent).find(node => ['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill'].includes(node.localName))
    if (!fill || fill.localName === 'noFill') return 'transparent'
    if (fill.localName === 'gradFill') {
        // OOXML gradient stops and clockwise angle:
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill.angle
        const stopNodes = children(child(fill, 'gsLst'), 'gs')
        const stops = stopNodes.map(stop => {
            const color = readColor(children(stop)[0], theme) || 'transparent'
            return `${color} ${Math.max(0, Math.min(numberAttr(stop, 'pos') / 1000, 100))}%`
        })
        if (stops.length < 2) {
            warnings.add('A gradient with fewer than two stops was ignored.')
            return readColor(children(stopNodes[0])[0], theme) || 'transparent'
        }
        const path = child(fill, 'path')
        if (path) {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathgradientfill
            if (child(path, 'fillToRect')) warnings.add('A path-gradient fillToRect focus rectangle is approximated by a centered radial gradient.')
            if (path.getAttribute('path') !== 'circle') warnings.add('A non-circular path gradient was approximated with a radial gradient.')
            return `radial-gradient(${path.getAttribute('path') === 'circle' ? 'circle' : 'ellipse'}, ${stops.join(', ')})`
        }
        const angle = numberAttr(child(fill, 'lin'), 'ang') / 60_000
        return `linear-gradient(${(90 + angle) % 360}deg, ${stops.join(', ')})`
    }
    if (fill.localName !== 'solidFill') {
        warnings.add(`The ${fill.localName.replace('Fill', '').toLowerCase()} fill was simplified.`)
        return 'transparent'
    }
    return readColor(children(fill)[0], theme) || 'transparent'
}

function lineDashStyle(line: Element | undefined, warnings: Set<string>): PptxElement['strokeDash'] {
    const preset = child(line, 'prstDash')?.getAttribute('val')
    if (preset) {
        if (!['solid', 'dot', 'sysDot', 'dash', 'sysDash', 'lgDash', 'dashDot', 'sysDashDot', 'lgDashDot', 'lgDashDotDot', 'sysDashDotDot'].includes(preset)) {
            warnings.add(`The unsupported ${preset} line dash was approximated with a dash.`)
            return 'dash'
        }
        if (!['solid', 'dot', 'dash'].includes(preset)) warnings.add(`The ${preset} line dash length was approximated with a browser built-in pattern.`)
        return preset
    }
    const custom = child(line, 'custDash')
    if (custom) {
        // ECMA-376 DrawingML defines d/sp as dash and space lengths relative to line width.
        // https://ecma-international.org/publications-and-standards/standards/ecma-376/
        const stops = children(custom, 'ds')
        const pattern = stops.flatMap(stop => {
            const ratio = (value: string | null) => {
                if (!value) return undefined
                const raw = value.trim()
                const amount = raw.endsWith('%') ? Number(raw.slice(0, -1)) / 100 : Number(raw) / 100_000
                return Number.isFinite(amount) && amount > 0 ? amount : undefined
            }
            const dash = ratio(stop.getAttribute('d'))
            const space = ratio(stop.getAttribute('sp'))
            return dash === undefined || space === undefined ? [] : [dash, space]
        })
        if (stops.length && pattern.length === stops.length * 2) return pattern
        warnings.add('A malformed custom line dash was approximated with a dash.')
        return 'dash'
    }
    return undefined
}

function lineCapStyle(line: Element | undefined, warnings: Set<string>): PptxElement['strokeCap'] {
    const value = line?.getAttribute('cap')
    if (!value) return undefined
    if (value === 'flat') return 'butt'
    if (value === 'rnd') return 'round'
    if (value === 'sq') return 'square'
    warnings.add(`The unsupported ${value} line cap was rendered as flat.`)
    return 'butt'
}

// DrawingML shape guides define 17 formulas; angle arguments use 60,000ths of a degree.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguide.formula?view=openxml-3.0.1
function customGeometryGuideResolver(geometry: Element | undefined, width: number, height: number, warnings: Set<string>) {
    const nodes = [...children(child(geometry, 'avLst'), 'gd'), ...children(child(geometry, 'gdLst'), 'gd')]
    if (nodes.length > 1024) warnings.add('Custom geometry guide lists with more than 1024 entries were truncated.')
    const guides = new Map(
        nodes.slice(0, 1024).flatMap(node => {
            const name = node.getAttribute('name'),
                formula = node.getAttribute('fmla')
            return name && formula ? [[name, formula] as const] : []
        })
    )
    const shortSide = Math.min(width, height),
        longSide = Math.max(width, height)
    const values = new Map<string, number | null>()
    const resolving = new Set<string>()
    const builtins = new Map<string, number>([
        ['w', width],
        ['h', height],
        ['l', 0],
        ['t', 0],
        ['r', width],
        ['b', height],
        ['hc', width / 2],
        ['vc', height / 2],
        ['ss', shortSide],
        ['ls', longSide],
        ['cd2', 10_800_000],
        ['cd3', 7_200_000],
        ['cd4', 5_400_000],
        ['cd6', 3_600_000],
        ['cd8', 2_700_000],
        ['cd16', 1_350_000],
        ['3cd4', 16_200_000],
        ['3cd8', 8_100_000],
        ['5cd8', 13_500_000],
        ['7cd8', 18_900_000]
    ])
    for (const [prefix, size] of [
        ['wd', width],
        ['hd', height],
        ['ssd', shortSide],
        ['lsd', longSide]
    ] as const) {
        for (const divisor of [2, 3, 4, 5, 6, 8, 10, 16, 32]) builtins.set(`${prefix}${divisor}`, size / divisor)
    }
    const argumentCounts: Record<string, number> = {'*/': 3, '+-': 3, '+/': 3, '?:': 3, abs: 1, at2: 2, cat2: 3, cos: 2, max: 2, min: 2, mod: 3, pin: 3, sat2: 3, sin: 2, sqrt: 1, tan: 2, val: 1}
    const numericToken = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i

    const resolve = (token: string, depth = 0): number | undefined => {
        if (!token || depth > 64) return undefined
        if (numericToken.test(token)) {
            const number = Number(token)
            return Number.isFinite(number) ? number : undefined
        }
        const builtin = builtins.get(token)
        if (builtin !== undefined) return builtin
        const cached = values.get(token)
        if (cached !== undefined) return cached ?? undefined
        const formula = guides.get(token)
        if (!formula || resolving.has(token)) {
            values.set(token, null)
            return undefined
        }

        resolving.add(token)
        let value: number | undefined
        try {
            const [operator, ...operands] = formula.trim().split(/\s+/)
            if (!operator || operands.length !== argumentCounts[operator]) return undefined
            const args = operands.map(operand => resolve(operand, depth + 1))
            if (args.some(argument => argument === undefined)) return undefined
            const x = args[0]!,
                y = args[1]!,
                z = args[2]!
            const radians = (angle: number) => ((angle / 60_000) * Math.PI) / 180
            const degrees = (angle: number) => ((angle * 180) / Math.PI) * 60_000
            switch (operator) {
                case '*/':
                    value = z === 0 ? undefined : (x * y) / z
                    break
                case '+-':
                    value = x + y - z
                    break
                case '+/':
                    value = z === 0 ? undefined : (x + y) / z
                    break
                case '?:':
                    value = x > 0 ? y : z
                    break
                case 'abs':
                    value = Math.abs(x)
                    break
                case 'at2':
                    value = degrees(Math.atan(y / x))
                    break
                case 'cat2':
                    value = x * Math.cos(Math.atan(z / y))
                    break
                case 'cos':
                    value = x * Math.cos(radians(y))
                    break
                case 'max':
                    value = Math.max(x, y)
                    break
                case 'min':
                    value = Math.min(x, y)
                    break
                case 'mod':
                    value = Math.hypot(x, y, z)
                    break
                case 'pin':
                    value = y < x ? x : y > z ? z : y
                    break
                case 'sat2':
                    value = x * Math.sin(Math.atan(z / y))
                    break
                case 'sin':
                    value = x * Math.sin(radians(y))
                    break
                case 'sqrt':
                    value = x < 0 ? undefined : Math.sqrt(x)
                    break
                case 'tan':
                    value = x * Math.tan(radians(y))
                    break
                case 'val':
                    value = x
                    break
            }
        } finally {
            resolving.delete(token)
        }
        const result = value !== undefined && Number.isFinite(value) ? value : undefined
        values.set(token, result ?? null)
        return result
    }

    return resolve
}

function parseCustomGeometryTextRect(geometry: Element | undefined, warnings: Set<string>): PptxElement['customTextRect'] {
    const rect = child(geometry, 'rect')
    if (!rect) return undefined
    // The custom-geometry text rectangle uses the shape's coordinate space and may refer to guides.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customgeometry.rectangle?view=openxml-3.0.1
    const path = children(child(geometry, 'pathLst'), 'path')[0]
    const width = numberAttr(path, 'w'),
        height = numberAttr(path, 'h')
    if (!(width > 0) || !(height > 0)) {
        warnings.add('A custom geometry text rectangle has no valid coordinate space; the shape bounds are used.')
        return undefined
    }
    const resolveGuide = customGeometryGuideResolver(geometry, width, height, warnings)
    const values = (['l', 't', 'r', 'b'] as const).map(name => {
        const value = rect.getAttribute(name)
        return value === null ? undefined : resolveGuide(value)
    })
    if (values.some(value => value === undefined)) {
        warnings.add('A custom geometry text rectangle contains a missing or unresolved guide; the shape bounds are used.')
        return undefined
    }
    const [left, top, right, bottom] = values as [number, number, number, number]
    if (!(right > left) || !(bottom > top)) {
        warnings.add('A custom geometry text rectangle has invalid bounds; the shape bounds are used.')
        return undefined
    }
    return {left: left / width, top: top / height, right: right / width, bottom: bottom / height}
}

// DrawingML custom paths define their own coordinate system and can contain multiple paths.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customgeometry?view=openxml-3.0.1
// DrawingML arcTo uses the current point as the angle-locked start on its ellipse and sweeps clockwise.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.arcto?view=openxml-3.0.1
function parseCustomGeometry(geometry: Element | undefined, warnings: Set<string>): PptxCustomPath[] | undefined {
    const nodes = children(child(geometry, 'pathLst'), 'path')
    if (!nodes.length) {
        warnings.add('A custom geometry without paths was rendered as a rectangle.')
        return undefined
    }
    if (nodes.length > 16) warnings.add('Custom geometries with more than 16 paths were truncated.')
    let unsupported = false
    let unresolvedGuide = false
    const paths = nodes.slice(0, 16).flatMap((path): PptxCustomPath[] => {
        const width = numberAttr(path, 'w'),
            height = numberAttr(path, 'h')
        if (width <= 0 || height <= 0) {
            unsupported = true
            return []
        }
        const commands = children(path)
        if (commands.length > 2048) {
            unsupported = true
            return []
        }
        const resolveGuide = customGeometryGuideResolver(geometry, width, height, warnings)
        const coordinate = (point: Element, key: 'x' | 'y', size: number) => {
            const value = point.getAttribute(key)
            if (value === null) return undefined
            const resolved = resolveGuide(value)
            if (resolved === undefined) {
                unresolvedGuide = true
                return undefined
            }
            return (resolved / size) * 1000
        }
        const d: string[] = []
        let hasMove = false
        let currentPoint: [number, number] | undefined
        let subpathStart: [number, number] | undefined
        for (const command of commands) {
            const points = children(command, 'pt')
            const coordinates = (point: Element): [number, number] | undefined => {
                const x = coordinate(point, 'x', width),
                    y = coordinate(point, 'y', height)
                return x === undefined || y === undefined ? undefined : [x, y]
            }
            const formatPoint = ([x, y]: [number, number]) => `${x} ${y}`
            if (command.localName === 'moveTo' || command.localName === 'lnTo') {
                const values = points.map(coordinates)
                if (!values.length || values.some(value => value === undefined) || (command.localName === 'moveTo' && values.length !== 1) || (command.localName === 'lnTo' && !hasMove)) {
                    unsupported = true
                    break
                }
                const validPoints = values as [number, number][]
                d.push(`${command.localName === 'moveTo' ? 'M' : 'L'} ${validPoints.map(formatPoint).join(' ')}`)
                currentPoint = validPoints.at(-1)
                if (command.localName === 'moveTo') {
                    hasMove = true
                    subpathStart = currentPoint
                }
            } else if (command.localName === 'cubicBezTo' || command.localName === 'quadBezTo') {
                const expected = command.localName === 'cubicBezTo' ? 3 : 2
                const values = points.map(coordinates)
                if (!hasMove || values.length !== expected || values.some(value => value === undefined)) {
                    unsupported = true
                    break
                }
                const validPoints = values as [number, number][]
                d.push(`${command.localName === 'cubicBezTo' ? 'C' : 'Q'} ${validPoints.map(formatPoint).join(' ')}`)
                currentPoint = validPoints.at(-1)
            } else if (command.localName === 'arcTo') {
                const values = ['wR', 'hR', 'stAng', 'swAng'].map(attribute => command.getAttribute(attribute))
                if (values.some(value => value === null)) {
                    unsupported = true
                    unresolvedGuide = true
                    break
                }
                const resolvedValues = values.map(value => resolveGuide(value || ''))
                const [radiusXValue, radiusYValue, startAngleValue, sweepAngleValue] = resolvedValues
                if (resolvedValues.some(value => value === undefined)) {
                    unsupported = true
                    unresolvedGuide = true
                    break
                }
                const radiusX = (radiusXValue! / width) * 1000
                const radiusY = (radiusYValue! / height) * 1000
                if (!hasMove || !currentPoint || !Number.isFinite(radiusX) || !Number.isFinite(radiusY) || !(radiusX > 0) || !(radiusY > 0) || !Number.isFinite(startAngleValue) || !Number.isFinite(sweepAngleValue) || sweepAngleValue! < 0 || sweepAngleValue! > 21_600_000) {
                    unsupported = true
                    break
                }
                const startAngle = (((startAngleValue! / 60_000) % 360) * Math.PI) / 180
                const sweepAngle = ((sweepAngleValue! / 60_000) * Math.PI) / 180
                const centerX = currentPoint[0] - radiusX * Math.cos(startAngle)
                const centerY = currentPoint[1] - radiusY * Math.sin(startAngle)
                const segmentCount = Math.ceil(Math.abs(sweepAngle) / Math.PI)
                for (let segment = 1; segment <= segmentCount; segment++) {
                    const angle = startAngle + (sweepAngle * segment) / segmentCount
                    currentPoint = [centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]
                    d.push(`A ${radiusX} ${radiusY} 0 0 ${sweepAngle >= 0 ? 1 : 0} ${formatPoint(currentPoint)}`)
                }
            } else if (command.localName === 'close') {
                if (!hasMove || !subpathStart) {
                    unsupported = true
                    break
                }
                d.push('Z')
                currentPoint = subpathStart
            } else {
                unsupported = true
                break
            }
        }
        if (!d.length || !hasMove) {
            unsupported = true
            return []
        }
        const fillMode = path.getAttribute('fill') || 'norm'
        if (!['norm', 'none'].includes(fillMode)) warnings.add(`The custom geometry ${fillMode} path fill mode was simplified to normal fill.`)
        return [{d: d.join(' '), fill: fillMode !== 'none', stroke: !['0', 'false'].includes((path.getAttribute('stroke') || 'true').toLowerCase())}]
    })
    if (unresolvedGuide) warnings.add('A custom geometry path references a missing, cyclic, or unsupported guide formula; dependent segments were omitted.')
    if (unsupported) warnings.add('Unsupported or malformed custom geometry path segments were omitted.')
    return paths.length ? paths : undefined
}

const textUnderlineStyles: Record<string, PptxRun['underlineStyle']> = {
    none: undefined, words: undefined, sng: undefined, dbl: 'double', heavy: undefined,
    dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashHeavy: 'dashed', dashLong: 'dashed', dashLongHeavy: 'dashed',
    dotDash: 'dashed', dotDashHeavy: 'dashed', dotDotDash: 'dashed', dotDotDashHeavy: 'dashed',
    wavy: 'wavy', wavyHeavy: 'wavy', wavyDbl: 'wavy',
}
const approximatedTextUnderlineValues = new Set(['words', 'heavy', 'dottedHeavy', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavyHeavy', 'wavyDbl'])

function textStyle(style: Element | undefined, theme: Record<string, string>, warnings: Set<string>): PptxTextStyle {
    const fill = child(style, 'solidFill')
    const line = child(style, 'ln')
    const lineFill = children(line).find(node => ['solidFill', 'noFill', 'gradFill', 'pattFill'].includes(node.localName))
    let outlineColor: string | undefined
    let outlineWidth: number | undefined
    if (line) {
        if (lineFill?.localName === 'noFill') {
            outlineColor = 'transparent'
            outlineWidth = 0
        } else if (lineFill && lineFill.localName !== 'solidFill') {
            warnings.add('Gradient and pattern text outlines are not supported.')
            outlineColor = 'transparent'
            outlineWidth = 0
        } else {
            // a:ln is valid on run properties. Solid outlines map directly to the browser text-stroke.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outline?view=openxml-3.0.1
            outlineColor = readColor(children(lineFill)[0], theme) || theme.dk1 || '#000000'
            outlineWidth = Math.max(numberAttr(line, 'w'), 9525)
            if (children(line).some(node => !['solidFill', 'noFill', 'gradFill', 'pattFill'].includes(node.localName)) || ['algn', 'cmpd', 'cap'].some(name => line.hasAttribute(name))) {
                warnings.add('Additional DrawingML text outline line styles are simplified to a solid stroke.')
            }
        }
    }
    const typeface = (script: 'latin' | 'ea' | 'cs') => {
        const value = child(style, script)?.getAttribute('typeface')?.trim()
        const token = value && /^\+m[jn]-(lt|ea|cs)$/.exec(value)
        if (token) {
            const family = token[0].slice(1, 3) === 'mj' ? 'major' : 'minor'
            const tokenScript = token[1] === 'lt' ? 'latin' : token[1]
            return theme[`font-${family}-${tokenScript}`]
        }
        return value || theme[`font-minor-${script}`]
    }
    // a:rPr/@spc is ST_TextSpacingPoint, stored in hundredths of a point.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/b6aeb5af-90ac-41dc-8964-f7ba99ed5181
    const rawLetterSpacing = style?.getAttribute('spc')
    const letterSpacing = rawLetterSpacing == null ? undefined : Number(rawLetterSpacing)
    const validLetterSpacing = letterSpacing !== undefined && Number.isInteger(letterSpacing) && letterSpacing >= -2_147_483_648 && letterSpacing <= 2_147_483_647
    if (rawLetterSpacing != null && !validLetterSpacing) warnings.add('A text run with malformed character spacing was rendered with default spacing.')
    // a:rPr/@baseline is ST_Percentage in thousandths of a percent; positive values raise the run.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/ff18a37e-9bd7-4338-9c37-1e285b5a5dd2
    const baseline = style?.getAttribute('baseline')?.trim()
    const baselineIsPercent = baseline?.endsWith('%')
    const baselineValue = baseline == null ? undefined : Number(baselineIsPercent ? baseline.slice(0, -1) : baseline)
    const validBaseline =
        baseline == null || (baselineIsPercent ? /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(baseline) && Number.isFinite(baselineValue) : /^[+-]?\d+$/.test(baseline) && baselineValue !== undefined && Number.isInteger(baselineValue) && baselineValue >= -2_147_483_648 && baselineValue <= 2_147_483_647)
    if (baseline != null && !validBaseline) warnings.add('A text run with malformed baseline shift was rendered on the default baseline.')
    const rawUnderline = style?.hasAttribute('u') ? style.getAttribute('u')! : undefined
    const validUnderline = rawUnderline === undefined || Object.prototype.hasOwnProperty.call(textUnderlineStyles, rawUnderline)
    if (rawUnderline !== undefined && !validUnderline) warnings.add(`A text run with unsupported underline value "${rawUnderline}" was rendered as a solid underline.`)
    else if (rawUnderline !== undefined && approximatedTextUnderlineValues.has(rawUnderline)) warnings.add(`The DrawingML underline value "${rawUnderline}" is approximated as CSS ${textUnderlineStyles[rawUnderline] || 'solid'}.`)
    const strike = style?.getAttribute('strike')
    const validStrike = strike == null || ['noStrike', 'sngStrike', 'dblStrike'].includes(strike)
    if (!validStrike) warnings.add(`A text run with unsupported strike value "${strike}" was rendered without strikethrough.`)
    return {
        fontFamily: typeface('latin'),
        fontFamilyEastAsia: typeface('ea'),
        fontFamilyComplexScript: typeface('cs'),
        fontSizePt: style?.hasAttribute('sz') ? numberAttr(style, 'sz') / 100 : undefined,
        letterSpacingPt: validLetterSpacing ? letterSpacing / 100 : undefined,
        baselineOffsetEm: validBaseline && baselineValue !== undefined ? baselineValue / (baselineIsPercent ? 100 : 100_000) : undefined,
        bold: style?.hasAttribute('b') ? style.getAttribute('b') === '1' : undefined,
        italic: style?.hasAttribute('i') ? style.getAttribute('i') === '1' : undefined,
        underline: rawUnderline === undefined ? undefined : rawUnderline !== 'none',
        underlineStyle: validUnderline && rawUnderline !== undefined ? textUnderlineStyles[rawUnderline] : undefined,
        lineThrough: validStrike && strike !== undefined ? strike !== 'noStrike' : undefined,
        lineThroughStyle: validStrike && strike !== undefined ? strike === 'dblStrike' ? 'double' : 'solid' : undefined,
        color: readColor(children(fill)[0], theme),
        outlineColor,
        outlineWidth,
        shadow: outerShadow(style, theme)
    }
}

function mergeTextStyles(...styles: PptxTextStyle[]): PptxTextStyle {
    const result: PptxTextStyle = {}
    for (const style of styles) {
        for (const [key, value] of Object.entries(style) as [keyof PptxTextStyle, PptxTextStyle[keyof PptxTextStyle]][]) {
            if (value !== undefined) Object.assign(result, {[key]: value})
        }
    }
    return result
}

function warnTextDecorationStyleConflict(style: PptxTextStyle, warnings: Set<string>): void {
    if (style.underline && style.lineThrough && (style.underlineStyle || 'solid') !== (style.lineThroughStyle || 'solid')) {
        warnings.add('A text run uses different underline and strikethrough styles; CSS applies one style to both and prioritizes the underline style.')
    }
}

function parseSpacing(properties: Element | undefined, name: string): PptxParagraph['lineSpacing'] {
    const spacing = child(properties, name)
    const points = child(spacing, 'spcPts')
    if (points) return {unit: 'emu', value: numberAttr(points, 'val') * 127}
    const percent = child(spacing, 'spcPct')
    if (percent) return {unit: 'ratio', value: numberAttr(percent, 'val') / 100_000}
    return undefined
}

function parseBullet(properties: Element | undefined, warnings: Set<string>): PptxParagraph['bullet'] {
    if (child(properties, 'buNone')) return undefined
    const character = child(properties, 'buChar')
    if (character) return {kind: 'char', char: character.getAttribute('char') || '•'}
    const automatic = child(properties, 'buAutoNum')
    if (automatic) {
        const type = automatic.getAttribute('type') || 'arabicPeriod'
        const supported = new Set([
            'arabicParenR',
            'arabicParenBoth',
            'arabicPeriod',
            'arabicPlain',
            'romanLcParenR',
            'romanLcParenBoth',
            'romanLcPeriod',
            'romanUcParenR',
            'romanUcParenBoth',
            'romanUcPeriod',
            'alphaLcParenR',
            'alphaLcParenBoth',
            'alphaLcPeriod',
            'alphaUcParenR',
            'alphaUcParenBoth',
            'alphaUcPeriod'
        ])
        if (!supported.has(type)) warnings.add(`The ${type} numbered bullet style is simplified to decimal numbering.`)
        return {kind: 'number', type, startAt: automatic.hasAttribute('startAt') ? numberAttr(automatic, 'startAt', 1) : undefined}
    }
    if (child(properties, 'buBlip')) warnings.add('Picture bullets are not rendered.')
    return undefined
}

function parseParagraphDefaults(properties: Element | undefined, theme: Record<string, string>, warnings: Set<string>): ParagraphDefaults {
    const values: ParagraphValues = {}
    const alignment = properties?.getAttribute('algn')
    if (alignment) {
        const alignments: Record<string, PptxParagraph['align']> = {l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', thaiDist: 'justify'}
        if (alignments[alignment]) values.align = alignments[alignment]
    }
    if (properties?.hasAttribute('lvl')) values.level = numberAttr(properties, 'lvl')
    if (properties?.hasAttribute('marL')) values.marginLeft = numberAttr(properties, 'marL')
    if (properties?.hasAttribute('marR')) values.marginRight = numberAttr(properties, 'marR')
    if (properties?.hasAttribute('indent')) values.indent = numberAttr(properties, 'indent')
    const spaceBefore = parseSpacing(properties, 'spcBef')
    const spaceAfter = parseSpacing(properties, 'spcAft')
    const lineSpacing = parseSpacing(properties, 'lnSpc')
    if (spaceBefore) values.spaceBefore = spaceBefore
    if (spaceAfter) values.spaceAfter = spaceAfter
    if (lineSpacing) values.lineSpacing = lineSpacing
    if (children(properties).some(item => ['buNone', 'buChar', 'buAutoNum', 'buBlip'].includes(item.localName))) {
        values.bullet = parseBullet(properties, warnings)
    }
    return {values, text: textStyle(child(properties, 'defRPr') || child(properties, 'endParaRPr'), theme, warnings)}
}

function mergeParagraphDefaults(...defaults: (ParagraphDefaults | undefined)[]): ParagraphDefaults {
    const values: ParagraphValues = {}
    let text: PptxTextStyle = {}
    for (const current of defaults) {
        if (!current) continue
        Object.assign(values, current.values)
        text = mergeTextStyles(text, current.text)
    }
    return {values, text}
}

function mergeParagraphDefaultsByLevel(...maps: (ParagraphDefaultsByLevel | undefined)[]): ParagraphDefaultsByLevel {
    const result: ParagraphDefaultsByLevel = {}
    for (let level = 0; level < 9; level++) {
        const defaults = maps.map(map => map?.[level]).filter((item): item is ParagraphDefaults => Boolean(item))
        if (defaults.length) result[level] = mergeParagraphDefaults(...defaults)
    }
    return result
}

function paragraphDefaultsByLevel(styleList: Element | undefined, theme: Record<string, string>, warnings: Set<string>): ParagraphDefaultsByLevel {
    const fallback = parseParagraphDefaults(child(styleList, 'defPPr'), theme, warnings)
    const result: ParagraphDefaultsByLevel = {}
    for (let level = 0; level < 9; level++) {
        result[level] = mergeParagraphDefaults(fallback, parseParagraphDefaults(child(styleList, `lvl${level + 1}pPr`), theme, warnings))
    }
    return result
}

function placeholderParagraphDefaults(shape: Element, theme: Record<string, string>, warnings: Set<string>): ParagraphDefaultsByLevel {
    const textBody = child(shape, 'txBody')
    const result = paragraphDefaultsByLevel(child(textBody, 'lstStyle'), theme, warnings)
    for (const paragraph of children(textBody, 'p')) {
        const properties = child(paragraph, 'pPr')
        const level = properties?.hasAttribute('lvl') ? numberAttr(properties, 'lvl') : 0
        const paragraphDefaults = parseParagraphDefaults(properties, theme, warnings)
        paragraphDefaults.text = mergeTextStyles(paragraphDefaults.text, textStyle(child(paragraph, 'endParaRPr'), theme, warnings))
        result[level] = mergeParagraphDefaults(result[level], paragraphDefaults)
    }
    return result
}

// Office uses txBody.lstStyle when the paragraph has no matching local list style.
// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/581b9df7-eebf-4f94-8961-8239684b3a9e
function parseText(
    shape: Element,
    theme: Record<string, string>,
    warnings: Set<string>,
    inheritedStyle: PptxTextStyle = {},
    inheritedParagraphStyles: ParagraphDefaultsByLevel = {}
): {
    paragraphs: PptxParagraph[]
    margins: PptxElement['margins']
    verticalAlign: PptxElement['verticalAlign']
    textWrap?: PptxElement['textWrap']
    textOrientation?: PptxTextOrientation
    textColumnCount?: number
    textColumnSpacing?: number
    textFontScale?: number
    textLineSpacingReduction?: number
    textAutoFit?: PptxElement['textAutoFit']
    textAutoFitDynamic?: boolean
} {
    const txBody = child(shape, 'txBody')
    if (!txBody) return {paragraphs: [], margins: {left: 0, right: 0, top: 0, bottom: 0}, verticalAlign: 'top'}
    const body = child(txBody, 'bodyPr')
    const rawColumnCount = body?.getAttribute('numCol')
    const columnCount = rawColumnCount === null || rawColumnCount === undefined ? undefined : Number(rawColumnCount)
    const textColumnCount = columnCount !== undefined && Number.isSafeInteger(columnCount) && columnCount >= 1 && columnCount <= 16 ? columnCount : undefined
    if (columnCount !== undefined && textColumnCount === undefined) warnings.add('An invalid a:bodyPr numCol was ignored.')
    const rawColumnSpacing = body?.getAttribute('spcCol')
    const columnSpacing = rawColumnSpacing === null || rawColumnSpacing === undefined ? undefined : Number(rawColumnSpacing)
    const textColumnSpacing = columnSpacing !== undefined && Number.isSafeInteger(columnSpacing) && columnSpacing >= 0 && columnSpacing <= 2_147_483_647 ? columnSpacing : undefined
    if (columnSpacing !== undefined && textColumnSpacing === undefined) warnings.add('An invalid a:bodyPr spcCol was ignored.')
    const normalAutoFit = child(body, 'normAutofit')
    const shapeAutoFit = child(body, 'spAutoFit')
    const rawFontScale = normalAutoFit?.getAttribute('fontScale') ?? null
    const rawLineSpaceReduction = normalAutoFit?.getAttribute('lnSpcReduction') ?? null
    const parsedFontScale = fixedPercentage(rawFontScale)
    const parsedLineSpaceReduction = fixedPercentage(rawLineSpaceReduction)
    if (shapeAutoFit) warnings.add('a:spAutoFit shape resizing uses browser font metrics; PowerPoint may produce different dimensions.')
    if (rawFontScale !== null && parsedFontScale === undefined) warnings.add('An invalid a:normAutofit fontScale was ignored.')
    if (rawLineSpaceReduction !== null && parsedLineSpaceReduction === undefined) warnings.add('An invalid a:normAutofit lnSpcReduction was ignored.')
    const listParagraphStyles = paragraphDefaultsByLevel(child(txBody, 'lstStyle'), theme, warnings)
    const paragraphs: PptxParagraph[] = children(txBody, 'p').map(paragraph => {
        const props = child(paragraph, 'pPr')
        const level = props?.hasAttribute('lvl') ? numberAttr(props, 'lvl') : 0
        const resolved = mergeParagraphDefaults(inheritedParagraphStyles[level], listParagraphStyles[level], parseParagraphDefaults(props, theme, warnings))
        const paragraphStyle = mergeTextStyles(inheritedStyle, resolved.text)
        const runs: PptxRun[] = []
        for (const run of children(paragraph)) {
            if (run.localName === 'br') {
                // DrawingML a:br is a forced line break between runs in one paragraph.
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.break
                runs.push({text: '\n', ...mergeTextStyles(paragraphStyle, textStyle(child(run, 'rPr'), theme, warnings))})
                continue
            }
            if (!['r', 'fld'].includes(run.localName)) continue
            const textNode = child(run, 't')
            if (!textNode) continue
            const style = child(run, 'rPr') || child(run, 'endParaRPr')
            runs.push({text: textNode.textContent || '', ...mergeTextStyles(paragraphStyle, textStyle(style, theme, warnings))})
        }
        if (!runs.length && paragraph.textContent?.trim()) runs.push({text: paragraph.textContent.trim(), ...paragraphStyle})
        return {
            ...resolved.values,
            align: resolved.values.align || 'left',
            runs,
            level: props?.hasAttribute('lvl') ? level : resolved.values.level
        }
    })
    const margin = (key: string, fallback: number) => numberAttr(body, key, fallback)
    const anchor = body?.getAttribute('anchor')
    const wrap = body?.getAttribute('wrap')
    const rawTextOrientation = body?.getAttribute('vert')
    const validTextOrientations: PptxTextOrientation[] = ['horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl']
    if (rawTextOrientation && !validTextOrientations.includes(rawTextOrientation as PptxTextOrientation)) {
        warnings.add('The unsupported text vertical mode "' + rawTextOrientation + '" was ignored.')
    } else if (rawTextOrientation === 'wordArtVert' || rawTextOrientation === 'mongolianVert') {
        warnings.add(`The text vertical mode "${rawTextOrientation}" uses CSS writing modes; glyph shaping may differ from Office.`)
    } else if (rawTextOrientation === 'wordArtVertRtl') {
        warnings.add('The text vertical mode "wordArtVertRtl" is not rendered yet; horizontal text is used.')
    }
    if (wrap && !['square', 'none'].includes(wrap)) warnings.add(`The unsupported text wrapping mode "${wrap}" was ignored.`)
    return {
        paragraphs,
        margins: {left: margin('lIns', 91_440), right: margin('rIns', 91_440), top: margin('tIns', 45_720), bottom: margin('bIns', 45_720)},
        verticalAlign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
        textWrap: wrap === 'square' || wrap === 'none' ? wrap : undefined,
        textOrientation: rawTextOrientation && validTextOrientations.includes(rawTextOrientation as PptxTextOrientation) ? (rawTextOrientation as PptxTextOrientation) : undefined,
        textColumnCount,
        textColumnSpacing,
        textFontScale: normalAutoFit ? (parsedFontScale ?? 1) : undefined,
        textLineSpacingReduction: normalAutoFit ? (parsedLineSpaceReduction ?? 0) : undefined,
        textAutoFit: shapeAutoFit ? 'shape' : normalAutoFit ? 'normal' : undefined,
        textAutoFitDynamic: Boolean(normalAutoFit && parsedFontScale === undefined)
    }
}

function shapeTransform(node: Element | undefined) {
    const xfrm = firstDescendant(node, 'xfrm')
    const offset = child(xfrm, 'off')
    const extent = child(xfrm, 'ext')
    return {
        x: numberAttr(offset, 'x'),
        y: numberAttr(offset, 'y'),
        width: numberAttr(extent, 'cx'),
        height: numberAttr(extent, 'cy'),
        rotation: numberAttr(xfrm, 'rot') / 60_000,
        flipH: xfrm?.getAttribute('flipH') === '1',
        flipV: xfrm?.getAttribute('flipV') === '1'
    }
}

function parsePlaceholderIdentity(shape: Element): PptxElement['placeholder'] {
    const placeholder = firstDescendant(shape, 'ph')
    if (!placeholder) return undefined
    return {type: placeholder.getAttribute('type') || 'obj', index: placeholder.getAttribute('idx') || undefined}
}

function parseSpeakerNotes(xml: XMLDocument | undefined, theme: Record<string, string>, warnings: Set<string>): string | undefined {
    const tree = firstDescendant(xml?.documentElement, 'spTree')
    const bodies = children(tree, 'sp').filter(shape => parsePlaceholderIdentity(shape)?.type === 'body')
    const text = bodies
        .flatMap(shape => parseText(shape, theme, warnings).paragraphs)
        .map(paragraph => paragraph.runs.map(run => run.text).join(''))
        .join('\n')
        .trim()
    return text || undefined
}

function placeholderStyleCategory(type: string): 'titleStyle' | 'bodyStyle' | 'otherStyle' {
    if (['title', 'ctrTitle'].includes(type)) return 'titleStyle'
    if (['dt', 'ftr', 'hdr', 'sldNum'].includes(type)) return 'otherStyle'
    return 'bodyStyle'
}

// Slide-master txStyles carries separate title, body, and other text defaults.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.textstyles?view=openxml-3.0.1
function parseMasterTextStyles(xml: XMLDocument | undefined, theme: Record<string, string>, warnings: Set<string>): Record<string, MasterTextStyle> {
    const styles = child(xml?.documentElement, 'txStyles')
    const result: Record<string, MasterTextStyle> = {}
    for (const name of ['titleStyle', 'bodyStyle', 'otherStyle']) {
        const style = child(styles, name)
        result[name] = {
            style: textStyle(child(style, 'defRPr'), theme, warnings),
            paragraphs: paragraphDefaultsByLevel(style, theme, warnings)
        }
    }
    return result
}

function matchPlaceholder(identity: PptxElement['placeholder'], candidates: ParsedPlaceholder[]): ParsedPlaceholder | undefined {
    if (!identity) return undefined
    if (identity.index) {
        const indexed = candidates.find(candidate => candidate.index === identity.index)
        if (indexed) return indexed
    }
    return candidates.find(candidate => candidate.type === identity.type)
}

function inheritPlaceholderProperties(element: PptxElement, node: Element, fallback?: ParsedPlaceholder): PptxElement {
    if (!fallback) return element
    const fallbackElement = fallback.element
    const properties = child(node, 'spPr')
    const transform = child(properties, 'xfrm')
    const offset = child(transform, 'off')
    const extent = child(transform, 'ext')
    const fillNames = ['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill']
    const hasFill = children(properties).some(item => fillNames.includes(item.localName))
    const line = child(properties, 'ln')
    const hasLineFill = children(line).some(item => fillNames.includes(item.localName))
    const hasEffectProperties = children(properties).some(item => ['effectLst', 'effectDag'].includes(item.localName))
    const bodyPr = child(child(node, 'txBody'), 'bodyPr')
    const hasTextAutoFit = Boolean(child(bodyPr, 'normAutofit') || child(bodyPr, 'spAutoFit') || child(bodyPr, 'noAutofit'))
    const fallbackMargins = fallbackElement.margins
    const margin = (name: 'lIns' | 'rIns' | 'tIns' | 'bIns', fallbackValue: number) => {
        const current = name === 'lIns' ? element.margins.left : name === 'rIns' ? element.margins.right : name === 'tIns' ? element.margins.top : element.margins.bottom
        return bodyPr?.hasAttribute(name) ? current : fallbackValue
    }
    const anchor = bodyPr?.getAttribute('anchor')
    const currentGeometry = child(properties, 'prstGeom') || child(properties, 'custGeom')
    return {
        ...element,
        x: offset?.hasAttribute('x') ? element.x : fallbackElement.x,
        y: offset?.hasAttribute('y') ? element.y : fallbackElement.y,
        width: extent?.hasAttribute('cx') ? element.width : fallbackElement.width,
        height: extent?.hasAttribute('cy') ? element.height : fallbackElement.height,
        rotation: transform?.hasAttribute('rot') ? element.rotation : fallbackElement.rotation,
        flipH: transform?.hasAttribute('flipH') ? element.flipH : fallbackElement.flipH,
        flipV: transform?.hasAttribute('flipV') ? element.flipV : fallbackElement.flipV,
        geometry: currentGeometry ? element.geometry : fallbackElement.geometry,
        customPaths: currentGeometry ? element.customPaths : fallbackElement.customPaths,
        customTextRect: currentGeometry ? element.customTextRect : fallbackElement.customTextRect,
        fill: hasFill ? element.fill : fallbackElement.fill,
        stroke: hasLineFill ? element.stroke : fallbackElement.stroke,
        strokeWidth: line?.hasAttribute('w') ? element.strokeWidth : fallbackElement.strokeWidth,
        strokeDash: child(line, 'prstDash') || child(line, 'custDash') ? element.strokeDash : fallbackElement.strokeDash,
        strokeCap: line?.hasAttribute('cap') ? element.strokeCap : fallbackElement.strokeCap,
        lineHeadEnd: child(line, 'headEnd') ? element.lineHeadEnd : fallbackElement.lineHeadEnd,
        lineTailEnd: child(line, 'tailEnd') ? element.lineTailEnd : fallbackElement.lineTailEnd,
        // A slide placeholder inherits visual properties from its matching layout/master placeholder.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/b3e45aed-2196-42bd-8b88-c98dba85ffe4
        shadow: hasEffectProperties ? element.shadow : fallbackElement.shadow,
        margins: {
            left: margin('lIns', fallbackMargins.left),
            right: margin('rIns', fallbackMargins.right),
            top: margin('tIns', fallbackMargins.top),
            bottom: margin('bIns', fallbackMargins.bottom)
        },
        verticalAlign: anchor ? element.verticalAlign : fallbackElement.verticalAlign,
        textWrap: ['square', 'none'].includes(bodyPr?.getAttribute('wrap') || '') ? element.textWrap : fallbackElement.textWrap,
        textOrientation: bodyPr?.hasAttribute('vert') ? element.textOrientation : fallbackElement.textOrientation,
        textColumnCount: bodyPr?.hasAttribute('numCol') ? element.textColumnCount : fallbackElement.textColumnCount,
        textColumnSpacing: bodyPr?.hasAttribute('spcCol') ? element.textColumnSpacing : fallbackElement.textColumnSpacing,
        textFontScale: hasTextAutoFit ? element.textFontScale : fallbackElement.textFontScale,
        textLineSpacingReduction: hasTextAutoFit ? element.textLineSpacingReduction : fallbackElement.textLineSpacingReduction,
        textAutoFit: hasTextAutoFit ? element.textAutoFit : fallbackElement.textAutoFit,
        textAutoFitDynamic: hasTextAutoFit ? element.textAutoFitDynamic : fallbackElement.textAutoFitDynamic
    }
}

function parsePlaceholderCandidates(xml: XMLDocument | undefined, theme: Record<string, string>, warnings: Set<string>, masterStyles: Record<string, MasterTextStyle>, fallbacks: ParsedPlaceholder[] = [], warnPlaceholders = false): ParsedPlaceholder[] {
    const tree = xml && firstDescendant(xml.documentElement, 'spTree')
    const candidates: ParsedPlaceholder[] = []
    for (const node of children(tree)) {
        if (node.localName !== 'sp') continue
        const identity = parsePlaceholderIdentity(node)
        if (!identity) continue
        const fallback = matchPlaceholder(identity, fallbacks)
        const masterStyle = masterStyles[placeholderStyleCategory(identity.type)]
        const paragraphStyles = mergeParagraphDefaultsByLevel(masterStyle?.paragraphs, fallback?.paragraphStyles, placeholderParagraphDefaults(node, theme, warnings))
        let element = parseShape(node, theme, warnings, masterStyle?.style || {}, paragraphStyles, warnPlaceholders)
        element = inheritPlaceholderProperties(element, node, fallback)
        candidates.push({node, type: identity.type, index: identity.index, element, paragraphStyles})
    }
    return candidates
}

function parseShape(shape: Element, theme: Record<string, string>, warnings: Set<string>, inheritedStyle: PptxTextStyle = {}, inheritedParagraphStyles: ParagraphDefaultsByLevel = {}, warnPlaceholder = true): PptxElement {
    const properties = child(shape, 'spPr')
    const customGeometry = child(properties, 'custGeom')
    const geometry = child(properties, 'prstGeom')?.getAttribute('prst') || (customGeometry ? 'custom' : 'rect')
    const customPaths = customGeometry ? parseCustomGeometry(customGeometry, warnings) : undefined
    const customTextRect = customGeometry ? parseCustomGeometryTextRect(customGeometry, warnings) : undefined
    const supportedGeometry = new Set(['rect', 'roundRect', 'ellipse', 'line', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'hexagon', 'chevron', 'custom'])
    if (!supportedGeometry.has(geometry)) warnings.add(`The ${geometry} shape is simplified to a rectangle.`)
    if (geometry === 'custom' && !customPaths) warnings.add('An unsupported custom geometry was rendered as a rectangle.')
    const placeholder = parsePlaceholderIdentity(shape)
    if (placeholder && warnPlaceholder) warnings.add('Placeholder properties are only partially inherited from the layout or master.')
    const transform = shapeTransform(properties)
    const tx = parseText(shape, theme, warnings, inheritedStyle, inheritedParagraphStyles)
    const nonVisual = firstDescendant(shape, 'cNvPr')
    const line = child(properties, 'ln')
    let fill = fillColor(properties, theme, warnings)
    let stroke = fillColor(line, theme, warnings)
    const strokeDash = lineDashStyle(line, warnings)
    const strokeCap = lineCapStyle(line, warnings)
    if (Array.isArray(strokeDash) && !customGeometry && geometry !== 'line') warnings.add('Custom dash ratios on preset shape outlines are approximated with CSS dashes.')
    // DrawingML head/tail end sizes are relative to the line width.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendpropertiestype?view=openxml-3.0.1
    const parseLineEnd = (side: 'headEnd' | 'tailEnd'): PptxLineEnd | undefined => {
        const node = child(line, side)
        if (!node) return undefined
        const type = node.getAttribute('type') || 'none'
        const types: PptxLineEnd['type'][] = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow']
        if (!types.includes(type as PptxLineEnd['type'])) {
            warnings.add(`An unsupported ${side} line ending was omitted.`)
            return undefined
        }
        const size = (attribute: 'w' | 'len'): PptxLineEnd['width'] => {
            const value = node.getAttribute(attribute) || 'med'
            if (value === 'sm' || value === 'med' || value === 'lg') return value
            warnings.add(`An invalid ${side} ${attribute === 'w' ? 'width' : 'length'} was rendered at medium size.`)
            return 'med'
        }
        return {type: type as PptxLineEnd['type'], width: size('w'), length: size('len')}
    }
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'Shape',
        kind: geometry === 'line' ? 'line' : 'shape',
        geometry,
        customPaths,
        customTextRect,
        ...transform,
        fill,
        stroke,
        strokeWidth: numberAttr(line, 'w'),
        strokeDash,
        strokeCap,
        lineHeadEnd: parseLineEnd('headEnd'),
        lineTailEnd: parseLineEnd('tailEnd'),
        shadow: outerShadow(properties, theme),
        ...tx,
        placeholder
    }
}

function parseTable(frame: Element, theme: Record<string, string>, warnings: Set<string>): PptxElement | undefined {
    const table = firstDescendant(frame, 'tbl')
    if (!table) return undefined
    const columns = children(child(table, 'tblGrid'), 'gridCol').map(column => numberAttr(column, 'w'))
    const rows = children(table, 'tr').map(row => ({
        height: numberAttr(row, 'h'),
        cells: children(row, 'tc').map((cell): PptxTableCell => {
            const properties = child(cell, 'tcPr')
            const text = parseText(cell, theme, warnings)
            if (text.textAutoFit === 'shape') warnings.add('a:spAutoFit on table cells is not rendered; original cell and row sizes are retained.')
            const border = (name: string) => {
                const line = child(properties, name)
                if (!line) return undefined
                return {color: fillColor(line, theme, warnings), width: numberAttr(line, 'w', 12_700)}
            }
            const anchor = properties?.getAttribute('anchor')
            return {
                paragraphs: text.paragraphs,
                textColumnCount: text.textColumnCount,
                textColumnSpacing: text.textColumnSpacing,
                textFontScale: text.textFontScale,
                textLineSpacingReduction: text.textLineSpacingReduction,
                textAutoFitDynamic: text.textAutoFitDynamic,
                fill: fillColor(properties, theme, warnings),
                margins: {
                    left: numberAttr(properties, 'marL', 91_440),
                    right: numberAttr(properties, 'marR', 91_440),
                    top: numberAttr(properties, 'marT', 45_720),
                    bottom: numberAttr(properties, 'marB', 45_720)
                },
                verticalAlign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
                colSpan: Math.max(1, numberAttr(cell, 'gridSpan', 1)),
                rowSpan: Math.max(1, numberAttr(cell, 'rowSpan', 1)),
                hidden: ['1', 'true'].includes(cell.getAttribute('hMerge') || '') || ['1', 'true'].includes(cell.getAttribute('vMerge') || ''),
                borders: {
                    top: border('lnT'),
                    right: border('lnR'),
                    bottom: border('lnB'),
                    left: border('lnL')
                }
            }
        })
    }))
    if (firstDescendant(table, 'tableStyleId')) warnings.add('Referenced table styles are simplified; direct cell formatting is preserved.')
    const transform = shapeTransform(frame)
    const nonVisual = firstDescendant(frame, 'cNvPr')
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'Table',
        kind: 'table',
        geometry: 'rect',
        ...transform,
        fill: 'transparent',
        stroke: 'transparent',
        strokeWidth: 0,
        paragraphs: [],
        margins: {left: 0, right: 0, top: 0, bottom: 0},
        verticalAlign: 'top',
        table: {columns, rows}
    }
}

async function parseSmartArtTextFallback(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>): Promise<PptxElement | undefined> {
    const graphicData = firstDescendant(frame, 'graphicData')
    if (!graphicData?.getAttribute('uri')?.endsWith('/diagram')) return undefined

    const dataId = namespacedAttr(child(graphicData, 'relIds'), 'dm') || ''
    const relation = rels.get(dataId)
    let paragraphs: PptxParagraph[] = []
    if (!relation || relation.external || !relation.type.endsWith('/diagramData')) {
        warnings.add('SmartArt diagram data is missing or external; its text could not be recovered.')
    } else {
        const part = zip.file(relation.target)
        if (!part) warnings.add('The SmartArt diagram data part is missing; its text could not be recovered.')
        else {
            try {
                const xml = parseXml(await part.async('string'), relation.target)
                const points = children(firstDescendant(xml.documentElement, 'ptLst'), 'pt').filter(point => point.getAttribute('type') === 'node')
                for (const point of points) {
                    const pointText = child(point, 't')
                    if (!pointText) continue
                    const shape = pointText.ownerDocument.createElementNS(PRESENTATION_NS, 'p:sp')
                    const body = pointText.ownerDocument.createElementNS(PRESENTATION_NS, 'p:txBody')
                    for (const node of Array.from(pointText.childNodes)) body.appendChild(node.cloneNode(true))
                    shape.appendChild(body)
                    paragraphs.push(...parseText(shape, theme, warnings).paragraphs)
                }
            } catch {
                warnings.add('The SmartArt diagram data contains invalid XML; its text could not be recovered.')
            }
        }
    }

    if (!paragraphs.some(paragraph => paragraph.runs.some(run => run.text.trim()))) {
        paragraphs = [{align: 'left', runs: [{text: 'SmartArt text unavailable', fontSizePt: 14}]}]
    }
    warnings.add('SmartArt is shown as a plain text list; node layout, hierarchy, connectors, and formatting are not reproduced.')
    const nonVisual = firstDescendant(frame, 'cNvPr')
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'SmartArt text fallback',
        kind: 'shape',
        geometry: 'rect',
        ...shapeTransform(frame),
        fill: '#f8fafc',
        stroke: '#94a3b8',
        strokeWidth: 12_700,
        paragraphs,
        margins: {left: 182_880, right: 182_880, top: 137_160, bottom: 137_160},
        verticalAlign: 'top',
        textWrap: 'square'
    }
}

async function parseOlePreview(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>): Promise<PptxElement | undefined> {
    const graphicData = firstDescendant(frame, 'graphicData')
    const ole = child(graphicData, 'oleObj')
    if (!ole && graphicData?.getAttribute('uri') !== OLE_GRAPHIC_DATA_URI) return undefined

    const preview = child(ole, 'pic')
    const previewBlip = preview && firstDescendant(preview, 'blip')
    const previewRelation = rels.get(namespacedAttr(previewBlip, 'embed') || '')
    if (preview && previewRelation && !previewRelation.external && zip.file(previewRelation.target)) {
        const picture = await parsePicture(preview, rels, zip, theme, warnings, urls, imageUrlCache)
        if (picture.imageUrl) {
            warnings.add('An OLE object is displayed from its embedded preview image; its embedded or linked content is not opened.')
            const nonVisual = firstDescendant(frame, 'cNvPr')
            return {
                ...picture,
                id: nonVisual?.getAttribute('id') || picture.id,
                name: nonVisual?.getAttribute('name') || picture.name,
                ...shapeTransform(frame)
            }
        }
    }

    warnings.add('An OLE object has no usable embedded preview image; a labeled placeholder is shown.')
    const nonVisual = firstDescendant(frame, 'cNvPr')
    const label = ole?.getAttribute('progId') || ole?.getAttribute('name') || nonVisual?.getAttribute('name') || 'Embedded object'
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'OLE object preview',
        kind: 'shape',
        geometry: 'rect',
        ...shapeTransform(frame),
        fill: '#f8fafc',
        stroke: '#94a3b8',
        strokeWidth: 12_700,
        paragraphs: [{align: 'center', runs: [{text: label, fontSizePt: 14, bold: true}]}],
        margins: {left: 182_880, right: 182_880, top: 137_160, bottom: 137_160},
        verticalAlign: 'middle',
        textWrap: 'square'
    }
}

// ChartPart contains a DrawingML chart; cached points are the last chart data saved in OOXML.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.chartpart?view=openxml-3.0.1
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberingcache?view=openxml-3.0.1
function chartCache(source: Element | undefined, warnings: Set<string>): Array<string | null> {
    if (!source) return []
    const cache = ['strCache', 'numCache', 'strLit', 'numLit'].map(name => firstDescendant(source, name)).find(Boolean)
    if (!cache) return []
    const points = descendants(cache, 'pt').flatMap(point => {
        const index = Number(point.getAttribute('idx'))
        const value = child(point, 'v')?.textContent ?? ''
        return Number.isSafeInteger(index) && index >= 0 ? [{index, value}] : []
    })
    const declaredCount = numberAttr(child(cache, 'ptCount'), 'val', points.length)
    const count = Math.min(128, Math.max(declaredCount, ...points.map(point => point.index + 1), 0))
    if (declaredCount > 128 || points.some(point => point.index >= 128)) warnings.add('Chart categories beyond 128 are omitted.')
    const values: Array<string | null> = Array(count).fill(null)
    for (const point of points) {
        if (point.index < count) values[point.index] = point.value
    }
    return values
}

function chartFillColor(properties: Element | undefined, theme: Record<string, string>, warnings: Set<string>, fallback: string): string {
    const color = fillColor(properties, theme, warnings)
    if (color.includes('gradient')) {
        const firstStop = firstDescendant(child(properties, 'gradFill'), 'gs')
        warnings.add('Chart gradient fills are approximated with their first color stop.')
        return readColor(children(firstStop)[0], theme) || fallback
    }
    return color === 'transparent' ? fallback : color
}

function chartBoolean(node: Element | undefined): boolean | undefined {
    if (!node) return undefined
    return !['0', 'false', 'off'].includes((node.getAttribute('val') || '1').toLowerCase())
}

function chartDataLabelFlag(dataLabels: Element | undefined, name: string): boolean | undefined {
    if (!dataLabels) return undefined
    if (chartBoolean(child(dataLabels, 'delete')) === true) return false
    return chartBoolean(child(dataLabels, name))
}

function chartDataLabelFlagOverrides(dataLabels: Element | undefined, name: string): Record<number, boolean> {
    const overrides: Record<number, boolean> = {}
    for (const label of children(dataLabels, 'dLbl')) {
        const index = numberAttr(child(label, 'idx'), 'val', -1)
        const enabled = chartDataLabelFlag(label, name)
        if (index >= 0 && enabled !== undefined) overrides[index] = enabled
    }
    return overrides
}

const chartDataLabelPositions: PptxChartDataLabelPosition[] = ['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't']

function chartDataLabelPosition(dataLabels: Element | undefined): PptxChartDataLabelPosition | undefined {
    const value = child(dataLabels, 'dLblPos')?.getAttribute('val')
    return chartDataLabelPositions.includes(value as PptxChartDataLabelPosition) ? (value as PptxChartDataLabelPosition) : undefined
}

function chartDataLabelPositionOverrides(dataLabels: Element | undefined): Record<number, PptxChartDataLabelPosition> {
    const overrides: Record<number, PptxChartDataLabelPosition> = {}
    for (const label of children(dataLabels, 'dLbl')) {
        const index = numberAttr(child(label, 'idx'), 'val', -1)
        const position = chartDataLabelPosition(label)
        if (index >= 0 && position) overrides[index] = position
    }
    return overrides
}

function chartDataLabelSeparator(dataLabels: Element | undefined): string | undefined {
    const separator = child(dataLabels, 'separator')
    return separator ? separator.textContent || '' : undefined
}

function chartDataLabelSeparatorOverrides(dataLabels: Element | undefined): Record<number, string> {
    const overrides: Record<number, string> = {}
    for (const label of children(dataLabels, 'dLbl')) {
        const index = numberAttr(child(label, 'idx'), 'val', -1)
        const separator = chartDataLabelSeparator(label)
        if (index >= 0 && separator !== undefined) overrides[index] = separator
    }
    return overrides
}

function chartNumberFormat(dataLabels: Element | undefined): string | undefined {
    const numberFormat = child(dataLabels, 'numFmt')
    if (!numberFormat || !['0', 'false'].includes((numberFormat.getAttribute('sourceLinked') || '1').toLowerCase())) return undefined
    return numberFormat.getAttribute('formatCode')?.trim() || undefined
}

function chartValueNumberFormatOverrides(dataLabels: Element | undefined): Record<number, string> {
    const overrides: Record<number, string> = {}
    for (const label of children(dataLabels, 'dLbl')) {
        const index = numberAttr(child(label, 'idx'), 'val', -1)
        const format = chartNumberFormat(label)
        if (index >= 0 && format) overrides[index] = format
    }
    return overrides
}

function chartAxisPosition(axis: Element | undefined, axisName: string, warnings: Set<string>): PptxChartAxisPosition | undefined {
    const position = child(axis, 'axPos')?.getAttribute('val')
    if (position === null || position === undefined) return undefined
    if (['b', 'l', 'r', 't'].includes(position)) return position as PptxChartAxisPosition
    warnings.add(`Unsupported chart ${axisName} axis position uses its default side.`)
    return undefined
}

function chartMajorTickMark(axis: Element | undefined, axisName: string, warnings: Set<string>): PptxChartTickMark | undefined {
    const node = child(axis, 'majorTickMark')
    if (!node) return undefined
    const value = node.getAttribute('val')
    const tickMark = value === null ? 'cross' : value
    if (['cross', 'in', 'none', 'out'].includes(tickMark)) return tickMark as PptxChartTickMark
    warnings.add(`Unsupported chart ${axisName} major tick mark is ignored.`)
    return undefined
}

function chartTickLabelPosition(axis: Element | undefined, axisName: string, warnings: Set<string>): PptxChartTickLabelPosition | undefined {
    const node = child(axis, 'tickLblPos')
    if (!node) return undefined
    const value = node.getAttribute('val')
    const position = value === null ? 'nextTo' : value
    if (['high', 'low', 'nextTo', 'none'].includes(position)) return position as PptxChartTickLabelPosition
    warnings.add(`Unsupported chart ${axisName} tick label position uses its default.`)
    return undefined
}

function chartAxisDeleted(axis: Element | undefined, axisName: string, warnings: Set<string>): boolean {
    const node = child(axis, 'delete')
    if (!node) return false
    const value = node.getAttribute('val')
    if (value === null || value === '1' || value === 'true') return true
    if (value === '0' || value === 'false') return false
    warnings.add(`Invalid chart ${axisName} axis delete value is ignored.`)
    return false
}

function chartAxisCrossing(axis: Element | undefined, axisName: string, warnings: Set<string>, categoryCount?: number): {crosses?: PptxChartAxisCrosses; crossesAt?: number} {
    const crossesValue = child(axis, 'crosses')?.getAttribute('val')
    let crosses: PptxChartAxisCrosses | undefined
    if (crossesValue) {
        if (['autoZero', 'max', 'min'].includes(crossesValue)) crosses = crossesValue as PptxChartAxisCrosses
        else warnings.add(`Unsupported chart ${axisName} axis crossing uses its default.`)
    }
    const crossesAtNode = child(axis, 'crossesAt')
    const rawCrossesAt = crossesAtNode?.getAttribute('val')
    if (rawCrossesAt === null || rawCrossesAt === undefined) return crosses ? {crosses} : {}
    const parsedCrossesAt = rawCrossesAt.trim() ? Number(rawCrossesAt) : NaN
    if (!Number.isFinite(parsedCrossesAt) || (axisName === 'category' && (!Number.isInteger(parsedCrossesAt) || parsedCrossesAt < 1))) {
        warnings.add(`Invalid chart ${axisName} axis crossing value is ignored.`)
        return crosses ? {crosses} : {}
    }
    if (axisName === 'category' && categoryCount !== undefined && parsedCrossesAt > categoryCount) {
        warnings.add('A chart category-axis crossing index exceeds the categories and is clamped to the last category.')
        return {crossesAt: categoryCount}
    }
    return {...(crosses ? {crosses} : {}), crossesAt: parsedCrossesAt}
}

function chartValueAxis(axis: Element | undefined, warnings: Set<string>): (Omit<NonNullable<PptxChart['valueAxis']>, 'min' | 'max'> & {min?: number; max?: number}) | undefined {
    if (!axis) return undefined
    const deleted = chartAxisDeleted(axis, 'value', warnings)
    const scaling = child(axis, 'scaling')
    if (child(scaling, 'logBase')) {
        warnings.add('Logarithmic chart axes are rendered with linear SVG scales.')
        return deleted ? {deleted: true} : undefined
    }
    const readBound = (name: 'min' | 'max') => {
        const node = child(scaling, name)
        if (!node) return undefined
        const raw = node.getAttribute('val')
        const value = raw?.trim() ? Number(raw) : NaN
        if (!Number.isFinite(value)) {
            warnings.add('Invalid chart value-axis bounds are ignored.')
            return undefined
        }
        return value
    }
    let min = readBound('min')
    let max = readBound('max')
    if (min !== undefined && max !== undefined && min >= max) {
        warnings.add('Invalid chart value-axis bounds are ignored.')
        min = undefined
        max = undefined
    }
    const majorUnitNode = child(axis, 'majorUnit')
    const majorUnit = majorUnitNode ? Number(majorUnitNode.getAttribute('val')) : undefined
    if (majorUnitNode && (!Number.isFinite(majorUnit) || majorUnit! <= 0)) warnings.add('Invalid chart value-axis major units are ignored.')
    const numberFormatNode = child(axis, 'numFmt')
    const sourceLinked = (numberFormatNode?.getAttribute('sourceLinked') || '1').toLowerCase()
    const formatCode = numberFormatNode?.getAttribute('formatCode')?.trim()
    const position = chartAxisPosition(axis, 'value', warnings)
    const crossing = chartAxisCrossing(axis, 'value', warnings)
    const majorTickMark = chartMajorTickMark(axis, 'value', warnings)
    const tickLabelPosition = chartTickLabelPosition(axis, 'value', warnings)
    let numberFormat: string | undefined
    if (formatCode && ['0', 'false'].includes(sourceLinked)) {
        if (isSupportedChartNumberFormat(formatCode)) numberFormat = formatCode
        else warnings.add('Unsupported chart number formats are rendered as raw values.')
    } else if (formatCode) {
        warnings.add('Source-linked chart number formats use raw cached values; embedded workbook number styles are not resolved.')
    }
    return {
        ...(min !== undefined ? {min} : {}),
        ...(max !== undefined ? {max} : {}),
        ...(Number.isFinite(majorUnit) && majorUnit! > 0 ? {majorUnit} : {}),
        ...(numberFormat ? {numberFormat} : {}),
        ...(child(scaling, 'orientation')?.getAttribute('val') === 'maxMin' ? {reverse: true} : {}),
        ...(deleted ? {deleted: true} : {}),
        ...(child(axis, 'majorGridlines') ? {majorGridlines: true} : {}),
        ...(position ? {position} : {}),
        ...(majorTickMark ? {majorTickMark} : {}),
        ...(tickLabelPosition ? {tickLabelPosition} : {}),
        ...crossing
    }
}

async function parseChart(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>): Promise<PptxElement | undefined> {
    const reference = firstDescendant(frame, 'chart')
    const relation = rels.get(namespacedAttr(reference, 'id') || '')
    if (!reference || !relation || relation.external) {
        warnings.add('A chart with a missing or external chart relationship is not rendered.')
        return undefined
    }
    const part = zip.file(relation.target)
    if (!part) {
        warnings.add('A chart relationship points to a missing chart part.')
        return undefined
    }
    const xml = parseXml(await part.async('string'), relation.target)
    const chart = firstDescendant(xml.documentElement, 'chart')
    const plotArea = firstDescendant(chart, 'plotArea')
    const supported = ['barChart', 'lineChart', 'pieChart', 'doughnutChart']
    const chartNodes = children(plotArea).filter(item => item.localName.endsWith('Chart'))
    const typeNode = chartNodes.find(item => supported.includes(item.localName))
    if (!typeNode) {
        warnings.add(`The ${chartNodes[0]?.localName || 'unknown'} chart type is not rendered.`)
        return undefined
    }
    if (chartNodes.length > 1) warnings.add('A mixed chart is simplified to its first supported chart type.')
    const type = typeNode.localName === 'barChart' ? 'bar' : typeNode.localName === 'lineChart' ? 'line' : typeNode.localName === 'doughnutChart' ? 'doughnut' : 'pie'
    const isPieLike = type === 'pie' || type === 'doughnut'
    const firstSliceAngleNode = child(typeNode, 'firstSliceAng')
    const firstSliceAngleValue = firstSliceAngleNode ? Number(firstSliceAngleNode.getAttribute('val')) : 0
    const firstSliceAngle = Number.isInteger(firstSliceAngleValue) && firstSliceAngleValue >= 0 && firstSliceAngleValue <= 360 ? firstSliceAngleValue : 0
    if (firstSliceAngleNode && firstSliceAngle !== firstSliceAngleValue) warnings.add('Invalid chart first-slice angle uses 0 degrees.')
    const holeSizeNode = type === 'doughnut' ? child(typeNode, 'holeSize') : undefined
    const holeSizeValue = holeSizeNode ? Number(holeSizeNode.getAttribute('val')) : 50
    const holeSize = Number.isInteger(holeSizeValue) && holeSizeValue >= 10 && holeSizeValue <= 90 ? holeSizeValue : 50
    if (holeSizeNode && holeSize !== holeSizeValue) warnings.add('Invalid doughnut hole size uses the 50 percent default.')
    const direction = type === 'bar' && child(typeNode, 'barDir')?.getAttribute('val') === 'bar' ? 'horizontal' : 'vertical'
    const grouping = child(typeNode, 'grouping')?.getAttribute('val')
    if (grouping && !['clustered', 'standard'].includes(grouping)) warnings.add('Stacked and percentage-stacked charts are rendered as grouped series.')
    if (!isPieLike && children(plotArea).some(item => ['catAx', 'dateAx', 'serAx'].includes(item.localName))) warnings.add('Chart date and series axes, category-axis number formatting, minor tick marks, and advanced styling are simplified.')
    const legendNode = child(chart, 'legend')
    const legendPositionValue = child(legendNode, 'legendPos')?.getAttribute('val') || 'r'
    const legendPosition = legendPositionValue === 'b' ? 'bottom' : legendPositionValue === 't' ? 'top' : legendPositionValue === 'l' ? 'left' : legendPositionValue === 'r' ? 'right' : legendPositionValue === 'tr' ? 'topRight' : undefined
    if (legendNode && !legendPosition) warnings.add('Unsupported chart legend position is rendered on the right.')
    const legendOverlay = chartBoolean(child(legendNode, 'overlay')) ?? false
    const legendEntries = children(legendNode, 'legendEntry')
    const hiddenLegendEntries: number[] = []
    for (const entry of legendEntries) {
        if (chartBoolean(child(entry, 'delete')) !== true) continue
        const index = Number(child(entry, 'idx')?.getAttribute('val'))
        if (Number.isSafeInteger(index) && index >= 0) hiddenLegendEntries.push(index)
        else warnings.add('A chart legend entry with an invalid index was ignored.')
    }
    const legendLayout = child(legendNode, 'layout')
    const manualLayout = child(legendLayout, 'manualLayout')
    let invalidManualLayout = false
    const manualAxis = (axis: 'x' | 'y' | 'w' | 'h'): {value: number; mode: 'edge' | 'factor'} | undefined => {
        const coordinate = child(manualLayout, axis)
        if (!coordinate) return undefined
        const value = Number(coordinate.getAttribute('val'))
        const modeValue = child(manualLayout, `${axis}Mode`)?.getAttribute('val') || 'factor'
        if (modeValue !== 'edge' && modeValue !== 'factor') {
            invalidManualLayout = true
            warnings.add('An unsupported chart legend manual layout was ignored.')
            return undefined
        }
        const minimum = modeValue === 'edge' || axis === 'w' || axis === 'h' ? 0 : -1
        if (!coordinate.hasAttribute('val') || !Number.isFinite(value) || value < minimum || value > 1) {
            invalidManualLayout = true
            warnings.add('An out-of-range chart legend manual layout was ignored.')
            return undefined
        }
        return {value, mode: modeValue}
    }
    const manualLegendX = manualAxis('x')
    const manualLegendY = manualAxis('y')
    const manualLegendWidth = manualAxis('w')
    const manualLegendHeight = manualAxis('h')
    const parsedManualLayout =
        manualLayout && !invalidManualLayout && (manualLegendX || manualLegendY || manualLegendWidth || manualLegendHeight)
            ? {
                  ...(manualLegendX ? {x: manualLegendX} : {}),
                  ...(manualLegendY ? {y: manualLegendY} : {}),
                  ...(manualLegendWidth ? {w: manualLegendWidth} : {}),
                  ...(manualLegendHeight ? {h: manualLegendHeight} : {})
              }
            : undefined
    if (manualLayout && (manualLegendWidth || manualLegendHeight)) warnings.add('Manual chart legend contents are clipped to the specified box; text wrapping is simplified.')
    const unsupportedLegendLayout = (legendLayout && !manualLayout) || (manualLayout && children(manualLayout).some(item => !['layoutTarget', 'x', 'xMode', 'y', 'yMode', 'w', 'wMode', 'h', 'hMode'].includes(item.localName)))
    if (legendNode && (unsupportedLegendLayout || children(legendNode).some(item => ['spPr', 'txPr'].includes(item.localName)) || legendEntries.some(entry => !!child(entry, 'txPr')))) warnings.add('Chart legend sizing, automatic layout, and styling are simplified.')
    const palette = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map(key => theme[key] || DEFAULT_THEME[key]!)
    const seriesNodes = children(typeNode, 'ser')
    const chartDataLabels = child(typeNode, 'dLbls')
    const chartShowValueLabels = chartDataLabelFlag(chartDataLabels, 'showVal') ?? false
    const chartShowPercentLabels = chartDataLabelFlag(chartDataLabels, 'showPercent') ?? false
    const chartShowCategoryNameLabels = chartDataLabelFlag(chartDataLabels, 'showCatName') ?? false
    const chartShowSeriesNameLabels = chartDataLabelFlag(chartDataLabels, 'showSerName') ?? false
    const chartLabelSeparator = chartDataLabelSeparator(chartDataLabels)
    const chartValueFormat = chartNumberFormat(chartDataLabels)
    const chartLabelPosition = chartDataLabelPosition(chartDataLabels)
    const dataLabelNodes = [chartDataLabels, ...seriesNodes.map(item => child(item, 'dLbls'))].filter((item): item is Element => !!item)
    const pointDataLabels = dataLabelNodes.flatMap(item => children(item, 'dLbl'))
    const allDataLabelSettings = [...dataLabelNodes, ...pointDataLabels]
    const numberFormatNodes = allDataLabelSettings.map(item => child(item, 'numFmt')).filter((item): item is Element => !!item)
    if (numberFormatNodes.some(item => !['0', 'false'].includes((item.getAttribute('sourceLinked') || '1').toLowerCase()))) {
        warnings.add('Source-linked chart number formats use raw cached values; embedded workbook number styles are not resolved.')
    }
    if (numberFormatNodes.some(item => ['0', 'false'].includes((item.getAttribute('sourceLinked') || '1').toLowerCase()) && item.getAttribute('formatCode') && !isSupportedChartNumberFormat(item.getAttribute('formatCode')!.trim()))) {
        warnings.add('Unsupported chart number formats are rendered as raw values.')
    }
    if (allDataLabelSettings.some(item => ['showLegendKey', 'showBubbleSize'].some(name => chartBoolean(child(item, name)) === true))) {
        warnings.add('Chart legend-key and bubble-size data labels are not rendered.')
    }
    if (!isPieLike && allDataLabelSettings.some(item => chartBoolean(child(item, 'showPercent')) === true)) {
        warnings.add('Percentage data labels are rendered only for pie and doughnut charts.')
    }
    const labelPositionSettings = allDataLabelSettings.flatMap(item => [item, ...children(item, 'dLbl')])
    if (
        labelPositionSettings.some(item => {
            const position = child(item, 'dLblPos')
            return !!position && !chartDataLabelPositions.includes(position.getAttribute('val') as PptxChartDataLabelPosition)
        })
    )
        warnings.add('Unsupported chart data label positions use the default placement.')
    if (allDataLabelSettings.some(item => ['txPr', 'spPr'].some(name => !!child(item, name)))) warnings.add('Chart data label styling is simplified.')
    if (seriesNodes.length > 12) warnings.add('Chart series beyond 12 are omitted.')
    const series = seriesNodes
        .slice(0, 12)
        .map((item, index): PptxChartSeries => {
            const sourceIndex = Number(child(item, 'idx')?.getAttribute('val'))
            const legendIndex = Number.isSafeInteger(sourceIndex) && sourceIndex >= 0 ? sourceIndex : index
            if (child(item, 'idx') && legendIndex !== sourceIndex) warnings.add('A chart series with an invalid index uses its source order.')
            const values = chartCache(child(item, 'val'), warnings).map(value => {
                if (value === null || value.trim() === '') return null
                const number = Number(value)
                if (!Number.isFinite(number)) warnings.add('Malformed chart data points are skipped.')
                return Number.isFinite(number) ? number : null
            })
            const pointColors: Record<number, string> = {}
            for (const point of children(item, 'dPt')) {
                const pointIndex = numberAttr(child(point, 'idx'), 'val', -1)
                if (pointIndex >= 0 && pointIndex < 128) pointColors[pointIndex] = chartFillColor(child(point, 'spPr'), theme, warnings, palette[pointIndex % palette.length]!)
            }
            const tx = child(item, 'tx')
            const name =
                descendants(tx, 't')
                    .map(value => value.textContent || '')
                    .join('') ||
                firstDescendant(tx, 'v')?.textContent ||
                `Series ${index + 1}`
            const dataLabels = child(item, 'dLbls')
            const numberFormat = chartNumberFormat(dataLabels) ?? chartValueFormat
            return {
                legendIndex,
                name,
                values,
                color: chartFillColor(child(item, 'spPr'), theme, warnings, palette[index % palette.length]!),
                pointColors,
                showValueLabels: chartDataLabelFlag(dataLabels, 'showVal') ?? chartShowValueLabels,
                valueLabelOverrides: {...chartDataLabelFlagOverrides(chartDataLabels, 'showVal'), ...chartDataLabelFlagOverrides(dataLabels, 'showVal')},
                dataLabelPosition: chartDataLabelPosition(dataLabels) ?? chartLabelPosition,
                dataLabelPositionOverrides: {...chartDataLabelPositionOverrides(chartDataLabels), ...chartDataLabelPositionOverrides(dataLabels)},
                showPercentLabels: chartDataLabelFlag(dataLabels, 'showPercent') ?? chartShowPercentLabels,
                percentLabelOverrides: {...chartDataLabelFlagOverrides(chartDataLabels, 'showPercent'), ...chartDataLabelFlagOverrides(dataLabels, 'showPercent')},
                showCategoryNameLabels: chartDataLabelFlag(dataLabels, 'showCatName') ?? chartShowCategoryNameLabels,
                categoryNameLabelOverrides: {...chartDataLabelFlagOverrides(chartDataLabels, 'showCatName'), ...chartDataLabelFlagOverrides(dataLabels, 'showCatName')},
                showSeriesNameLabels: chartDataLabelFlag(dataLabels, 'showSerName') ?? chartShowSeriesNameLabels,
                seriesNameLabelOverrides: {...chartDataLabelFlagOverrides(chartDataLabels, 'showSerName'), ...chartDataLabelFlagOverrides(dataLabels, 'showSerName')},
                labelSeparator: chartDataLabelSeparator(dataLabels) ?? chartLabelSeparator,
                labelSeparatorOverrides: {...chartDataLabelSeparatorOverrides(chartDataLabels), ...chartDataLabelSeparatorOverrides(dataLabels)},
                ...(numberFormat ? {numberFormat} : {}),
                valueNumberFormatOverrides: {...chartValueNumberFormatOverrides(chartDataLabels), ...chartValueNumberFormatOverrides(dataLabels)}
            }
        })
        .filter(item => item.values.some(value => value !== null))
    if (!series.length) {
        warnings.add('A chart without cached series values cannot be rendered.')
        return undefined
    }
    if (isPieLike && series.length > 1) warnings.add('Pie and doughnut chart series beyond the first are omitted.')
    const parsedValueAxis = isPieLike ? undefined : chartValueAxis(children(plotArea, 'valAx')[0], warnings)
    const categoryAxis = children(plotArea, 'catAx')[0]
    let categoryAxisPosition = isPieLike ? undefined : chartAxisPosition(categoryAxis, 'category', warnings)
    const categoryAxisMajorTickMark = isPieLike ? undefined : chartMajorTickMark(categoryAxis, 'category', warnings)
    const categoryAxisTickLabelPosition = isPieLike ? undefined : chartTickLabelPosition(categoryAxis, 'category', warnings)
    const categoryAxisDeleted = !isPieLike && chartAxisDeleted(categoryAxis, 'category', warnings)
    if (direction === 'horizontal' && parsedValueAxis?.position && !['b', 't'].includes(parsedValueAxis.position)) {
        warnings.add('A horizontal chart value axis must be positioned on the top or bottom; its default side is used.')
        delete parsedValueAxis.position
    }
    if (direction === 'vertical' && parsedValueAxis?.position && !['l', 'r'].includes(parsedValueAxis.position)) {
        warnings.add('A vertical chart value axis must be positioned on the left or right; its default side is used.')
        delete parsedValueAxis.position
    }
    if (categoryAxisPosition && (direction === 'horizontal' ? !['l', 'r'].includes(categoryAxisPosition) : !['b', 't'].includes(categoryAxisPosition))) {
        warnings.add('The chart category axis position does not match its chart direction; its default side is used.')
        categoryAxisPosition = undefined
    }
    if (!isPieLike && children(plotArea, 'valAx').length > 1) warnings.add('Secondary chart value axes are omitted.')
    if (!isPieLike && children(plotArea, 'catAx').length > 1) warnings.add('Secondary chart category axes are omitted.')
    let valueAxis: PptxChart['valueAxis']
    if (parsedValueAxis) {
        const values = series.flatMap(item => item.values.filter((value): value is number => value !== null))
        const dataMin = Math.min(0, ...values)
        const dataMax = Math.max(0, ...values)
        let min = parsedValueAxis.min ?? dataMin
        let max = parsedValueAxis.max ?? (dataMax === dataMin ? dataMin + 1 : dataMax)
        if (min >= max) {
            warnings.add('Chart value-axis bounds conflict with the data range; data-driven bounds are used.')
            min = dataMin
            max = dataMax === dataMin ? dataMin + 1 : dataMax
            delete parsedValueAxis.min
            delete parsedValueAxis.max
        }
        if (parsedValueAxis.crossesAt !== undefined && (parsedValueAxis.crossesAt < min || parsedValueAxis.crossesAt > max)) {
            warnings.add('A chart value-axis crossing value is outside the displayed range and is clamped to the nearest bound.')
        }
        if (parsedValueAxis.majorUnit && Math.floor((max - min) / parsedValueAxis.majorUnit + 1e-9) > 100) {
            warnings.add('Chart value-axis major units produce too many ticks and are ignored.')
            delete parsedValueAxis.majorUnit
        }
        valueAxis = {...parsedValueAxis, min, max}
    }
    const categories = chartCache(child(seriesNodes[0], 'cat'), warnings)
    const pointCount = Math.min(128, Math.max(categories.length, ...series.map(item => item.values.length), 0))
    const labels = Array.from({length: pointCount}, (_, index) => categories[index] || String(index + 1))
    const categoryAxisCrossing: ReturnType<typeof chartAxisCrossing> = isPieLike ? {} : chartAxisCrossing(categoryAxis, 'category', warnings, Math.max(labels.length, 1))
    if (labels.length !== categories.length) warnings.add('Missing chart category labels are shown as point numbers.')
    if (isPieLike && labels.length > 12) warnings.add('Pie and doughnut chart legend entries beyond 12 are omitted.')
    if (isPieLike && series[0]!.values.some(value => value !== null && value < 0)) warnings.add('Negative pie and doughnut chart values are omitted.')
    const titleNode = firstDescendant(chart, 'title')
    const title =
        descendants(titleNode, 't')
            .map(value => value.textContent || '')
            .join('')
            .trim() || undefined
    const transform = shapeTransform(frame)
    const nonVisual = firstDescendant(frame, 'cNvPr')
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'Chart',
        kind: 'chart',
        geometry: 'rect',
        ...transform,
        fill: 'transparent',
        stroke: 'transparent',
        strokeWidth: 0,
        paragraphs: [],
        margins: {left: 0, right: 0, top: 0, bottom: 0},
        verticalAlign: 'top',
        chart: {
            type,
            direction,
            ...(type === 'doughnut' ? {holeSize} : {}),
            ...(firstSliceAngleNode ? {firstSliceAngle} : {}),
            title,
            categories: labels,
            palette,
            series,
            ...(chartLabelSeparator !== undefined ? {labelSeparator: chartLabelSeparator} : {}),
            ...(chartValueFormat ? {numberFormat: chartValueFormat} : {}),
            ...(valueAxis ? {valueAxis} : {}),
            ...(categoryAxisPosition ? {categoryAxisPosition} : {}),
            ...(categoryAxisDeleted ? {categoryAxisDeleted: true} : {}),
            ...(categoryAxisMajorTickMark ? {categoryAxisMajorTickMark} : {}),
            ...(categoryAxisTickLabelPosition ? {categoryAxisTickLabelPosition} : {}),
            ...(categoryAxisCrossing.crosses ? {categoryAxisCrosses: categoryAxisCrossing.crosses} : {}),
            ...(categoryAxisCrossing.crossesAt !== undefined ? {categoryAxisCrossesAt: categoryAxisCrossing.crossesAt} : {}),
            ...(legendNode
                ? {
                      legend: {
                          position: legendPosition || 'right',
                          ...(hiddenLegendEntries.length ? {hiddenEntries: hiddenLegendEntries} : {}),
                          ...(legendOverlay ? {overlay: true} : {}),
                          ...(parsedManualLayout ? {manualLayout: parsedManualLayout} : {})
                      }
                  }
                : {}),
            ...(chartLabelPosition ? {dataLabelPosition: chartLabelPosition} : {})
        }
    }
}

function parseTheme(xml: XMLDocument | undefined): Record<string, string> {
    const result = {...DEFAULT_THEME}
    const scheme = xml && firstDescendant(xml.documentElement, 'clrScheme')
    for (const slot of children(scheme)) {
        const color = children(slot)[0]
        if (!color) continue
        const value = color.getAttribute(color.localName === 'sysClr' ? 'lastClr' : 'val')
        if (value) result[slot.localName] = value.startsWith('#') ? value : `#${value}`
    }
    const fonts = xml && firstDescendant(xml.documentElement, 'fontScheme')
    for (const role of ['major', 'minor']) {
        const fontSet = child(fonts, `${role}Font`)
        for (const script of ['latin', 'ea', 'cs']) {
            const typeface = child(fontSet, script)?.getAttribute('typeface')?.trim()
            if (typeface) result[`font-${role}-${script}`] = typeface
        }
    }
    return result
}

function imageMime(path: string): string {
    const ext = path.split('.').at(-1)?.toLowerCase()
    return ({png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp'} as Record<string, string>)[ext || ''] || 'application/octet-stream'
}

async function imageUrlFor(target: string | undefined, zip: JSZip, urls: string[], cache: Map<string, string>): Promise<string | undefined> {
    if (!target) return undefined
    const cached = cache.get(target)
    if (cached) return cached
    const image = zip.file(target)
    if (!image) return undefined
    const blob = await image.async('blob')
    const url = URL.createObjectURL(new Blob([blob], {type: imageMime(target)}))
    cache.set(target, url)
    urls.push(url)
    return url
}

function mediaMime(path: string): string | undefined {
    const ext = path.split('.').at(-1)?.toLowerCase()
    return (
        {
            mp4: 'video/mp4',
            m4v: 'video/mp4',
            webm: 'video/webm',
            ogv: 'video/ogg',
            mov: 'video/quicktime',
            mpg: 'video/mpeg',
            mpeg: 'video/mpeg',
            wmv: 'video/x-ms-wmv',
            avi: 'video/x-msvideo',
            mp3: 'audio/mpeg',
            m4a: 'audio/mp4',
            aac: 'audio/aac',
            wav: 'audio/wav',
            wave: 'audio/wav',
            ogg: 'audio/ogg',
            oga: 'audio/ogg',
            opus: 'audio/ogg',
            flac: 'audio/flac',
            wma: 'audio/x-ms-wma'
        } as Record<string, string>
    )[ext || '']
}

function mediaTimeOffsetMs(value: string | null): number | undefined {
    // MS-PPTX ST_UniversalTimeOffset: h/min/s/ms/µs/ns suffixes; an omitted unit means milliseconds.
    if (value === null) return 0
    const match = /^(\d+(?:\.\d+)?)(ms|ns|µs|s|min|h)?$/.exec(value.trim())
    if (!match) return undefined
    const multiplier = ({h: 3_600_000, min: 60_000, s: 1_000, ms: 1, µs: 0.001, ns: 0.000001} as Record<string, number>)[match[2] || 'ms']!
    const milliseconds = Number(match[1]) * multiplier
    return Number.isFinite(milliseconds) && milliseconds >= 0 ? milliseconds : undefined
}

function selectedTransition(slideRoot: Element, warnings: Set<string>): {transition: Element; durationSource: Element} | undefined {
    let transition = child(slideRoot, 'transition')
    if (transition) return {transition, durationSource: transition}
    const alternate = child(slideRoot, 'AlternateContent')
    const choices = children(alternate, 'Choice')
    const choiceTransitions = choices.map(choice => ({choice, transition: child(choice, 'transition')})).filter((candidate): candidate is {choice: Element; transition: Element} => Boolean(candidate.transition))
    const supportedChoices = choiceTransitions.filter(({choice}) =>
        (choice.getAttribute('Requires') || '')
            .trim()
            .split(/\s+/)
            .filter(Boolean)
            .every(prefix => SUPPORTED_MC_REQUIRES_NAMESPACES.has(choice.lookupNamespaceURI(prefix) || ''))
    )
    transition = supportedChoices
        .map(candidate => candidate.transition)
        .find(candidate => {
            const effect = children(candidate).find(node => !['extLst', 'sndAc', 'sound'].includes(node.localName))
            return !!effect && SUPPORTED_TRANSITIONS.has(effect.localName)
        })
    if (transition) return {transition, durationSource: transition}
    transition = child(child(alternate, 'Fallback'), 'transition')
    const durationSource = choiceTransitions.map(candidate => candidate.transition).find(candidate => candidate.hasAttributeNS(POWERPOINT_2010_NS, 'dur')) || transition
    if (choices.length && transition) warnings.add('An unsupported AlternateContent slide transition choice was skipped in favor of its fallback.')
    return transition ? {transition, durationSource: durationSource || transition} : undefined
}

function parseSlideAdvance(slideRoot: Element, warnings: Set<string>): Pick<PptxSlide, 'advanceOnClick' | 'advanceAfterMs'> {
    const transition = selectedTransition(slideRoot, warnings)?.transition
    if (!transition) return {}
    const advanceOnClick = transition.getAttribute('advClick')?.toLowerCase()
    if (advanceOnClick !== undefined && !['1', '0', 'true', 'false'].includes(advanceOnClick)) {
        warnings.add('The slide advance-on-click value is invalid and was ignored.')
    }
    const rawAdvanceAfter = transition.getAttribute('advTm')
    const advanceAfterMs = rawAdvanceAfter !== null ? Number(rawAdvanceAfter) : undefined
    if (rawAdvanceAfter !== null && (!/^\d+$/.test(rawAdvanceAfter) || !Number.isSafeInteger(advanceAfterMs) || advanceAfterMs! > 0xffff_ffff)) {
        warnings.add('The slide auto-advance time is invalid and was ignored.')
    }
    return {
        advanceOnClick: advanceOnClick === '0' || advanceOnClick === 'false' ? false : undefined,
        advanceAfterMs: rawAdvanceAfter !== null && /^\d+$/.test(rawAdvanceAfter) && Number.isSafeInteger(advanceAfterMs) && advanceAfterMs! <= 0xffff_ffff ? advanceAfterMs : undefined
    }
}

function parseTransition(slideRoot: Element, warnings: Set<string>): PptxSlide['transition'] {
    const selection = selectedTransition(slideRoot, warnings)
    if (!selection) return undefined
    const {transition, durationSource} = selection
    const effectNode = children(transition).find(node => !['extLst', 'sndAc', 'sound'].includes(node.localName))
    if (!effectNode) return undefined
    const effectName = effectNode.localName
    const throughBlackValue = effectNode.getAttribute('thruBlk')?.toLowerCase()
    const throughBlack = ['fade', 'cut', 'reveal'].includes(effectName) && ['1', 'true'].includes(throughBlackValue || '')
    if (throughBlackValue !== undefined && ['fade', 'cut', 'reveal'].includes(effectName) && !['0', '1', 'false', 'true'].includes(throughBlackValue)) warnings.add(`The ${effectName} transition has an invalid thruBlk value.`)
    const rawPresetName = effectName === 'prstTrans' ? effectNode.getAttribute('prst') : null
    const presetName = rawPresetName && SUPPORTED_PRESET_TRANSITIONS.has(rawPresetName) ? (rawPresetName as NonNullable<PptxSlide['transition']>['presetName']) : undefined
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.fadetransition?view=openxml-3.0.1
    const effect: NonNullable<PptxSlide['transition']>['effect'] =
        effectName === 'prstTrans'
            ? presetName
                ? 'preset'
                : 'fade'
            : [
                    'fade',
                    'push',
                    'wipe',
                    'zoom',
                    'cover',
                    'pull',
                    'split',
                    'circle',
                    'diamond',
                    'plus',
                    'wedge',
                    'newsflash',
                    'flash',
                    'doors',
                    'window',
                    'prism',
                    'pan',
                    'vortex',
                    'ferris',
                    'gallery',
                    'conveyor',
                    'ripple',
                    'glitter',
                    'shred',
                    'flythrough',
                    'warp',
                    'flip',
                    'switch',
                    'reveal',
                    'morph',
                    'random',
                    'cut',
                    'honeycomb',
                    'blinds',
                    'checker',
                    'comb',
                    'dissolve',
                    'wheel',
                    'wheelReverse',
                    'randomBar',
                    'strips'
                ].includes(effectName)
              ? (effectName as NonNullable<PptxSlide['transition']>['effect'])
              : 'fade'
    if (effectName === 'prstTrans' && !presetName) warnings.add(`The p15:prstTrans preset "${rawPresetName || '(missing)'}" is invalid; the transition falls back to fade.`)
    else if (effect !== effectName && effectName !== 'prstTrans') warnings.add(`The ${effectName} slide transition falls back to fade.`)
    const speed = transition.getAttribute('spd')
    // Office stores the extension duration on mc:Choice even when this player renders mc:Fallback.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/99b95b35-568a-4652-9cba-df3d1175952f
    const durationValue = (durationSource || transition).getAttributeNS(POWERPOINT_2010_NS, 'dur')
    const explicitDuration = durationValue === null ? undefined : mediaTimeOffsetMs(durationValue)
    if (durationValue !== null && explicitDuration === undefined) warnings.add('An invalid p14 slide transition duration was ignored.')
    // MS-PPTX defines p14:wheelReverse with the same CT_WheelTransition type and spokes attribute as wheel.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/76223734-7aa7-4053-b666-acd5f73d1d9a
    const rawSpokes = effect === 'wheel' || effect === 'wheelReverse' ? Number(effectNode?.getAttribute('spokes') || 4) : undefined
    const spokes = rawSpokes === undefined ? undefined : [1, 2, 3, 4, 8].reduce((nearest, value) => (Math.abs(value - rawSpokes) < Math.abs(nearest - rawSpokes) ? value : nearest), 4)
    if (rawSpokes !== undefined && (!Number.isFinite(rawSpokes) || rawSpokes < 1 || ![1, 2, 3, 4, 8].includes(rawSpokes))) warnings.add(`The wheel transition spoke count was normalized to ${spokes}.`)
    const rawDirection = effectNode?.getAttribute('dir')
    const stripsDirections = ['ld', 'lu', 'rd', 'ru']
    const sideDirections = ['l', 'r', 'u', 'd', 'ld', 'lu', 'rd', 'ru']
    const vortexDirections = ['l', 'r', 'u', 'd']
    const leftRightDirections = ['l', 'r']
    const rippleDirections = ['center', 'lu', 'ru', 'ld', 'rd']
    // MS-PPTX CT_GlitterTransition uses side direction and diamond/hexagon pattern.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/93a73ee5-6799-478f-8c1f-3db23726e10e
    const glitterPatterns = ['diamond', 'hexagon']
    const direction =
        effect === 'strips'
            ? stripsDirections.includes(rawDirection || 'ld')
                ? rawDirection || 'ld'
                : 'ld'
            : effect === 'shred'
              ? ['in', 'out'].includes(rawDirection || 'in')
                  ? rawDirection || 'in'
                  : 'in'
              : effect === 'flythrough'
                ? ['in', 'out'].includes(rawDirection || 'in')
                    ? rawDirection || 'in'
                    : 'in'
                : effect === 'warp'
                  ? ['in', 'out'].includes(rawDirection || 'out')
                      ? rawDirection || 'out'
                      : 'out'
                  : effect === 'vortex'
                    ? vortexDirections.includes(rawDirection || 'l')
                        ? rawDirection || 'l'
                        : 'l'
                    : ['prism', 'pan', 'vortex', 'glitter'].includes(effect)
                      ? sideDirections.includes(rawDirection || 'l')
                          ? rawDirection || 'l'
                          : 'l'
                      : ['flip', 'switch', 'reveal', 'ferris', 'gallery', 'conveyor'].includes(effect)
                        ? leftRightDirections.includes(rawDirection || 'l')
                            ? rawDirection || 'l'
                            : 'l'
                        : effect === 'ripple'
                          ? rippleDirections.includes(rawDirection || 'center')
                              ? rawDirection || 'center'
                              : 'center'
                          : rawDirection || (effect === 'zoom' || effect === 'split' ? 'in' : 'l')
    if (effect === 'strips' && rawDirection && !stripsDirections.includes(rawDirection)) warnings.add(`The strips transition direction "${rawDirection}" was normalized to ld.`)
    if (effect === 'shred' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The shred transition direction "${rawDirection}" was normalized to in.`)
    if (effect === 'flythrough' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The flythrough transition direction "${rawDirection}" was normalized to in.`)
    if (effect === 'warp' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The warp transition direction "${rawDirection}" was normalized to out.`)
    if (effect === 'vortex' && rawDirection && !vortexDirections.includes(rawDirection)) warnings.add(`The vortex transition direction "${rawDirection}" was normalized to l.`)
    if (['prism', 'pan'].includes(effect) && rawDirection && !sideDirections.includes(rawDirection)) warnings.add(`The ${effect} transition direction "${rawDirection}" was normalized to l.`)
    if (effect === 'glitter' && rawDirection && !sideDirections.includes(rawDirection)) warnings.add(`The glitter transition direction "${rawDirection}" was normalized to l.`)
    if (['flip', 'switch', 'reveal', 'ferris', 'gallery', 'conveyor'].includes(effect) && rawDirection && !leftRightDirections.includes(rawDirection)) warnings.add(`The ${effect} transition direction "${rawDirection}" was normalized to l.`)
    if (effect === 'ripple' && rawDirection && !rippleDirections.includes(rawDirection)) warnings.add(`The ripple transition direction "${rawDirection}" was normalized to center.`)
    const rawPattern = effect === 'glitter' ? effectNode.getAttribute('pattern') : null
    const pattern = effect === 'glitter' ? (rawPattern === 'hexagon' ? 'hexagon' : 'diamond') : undefined
    if (effect === 'glitter' && rawPattern && !glitterPatterns.includes(rawPattern)) warnings.add(`The glitter transition pattern "${rawPattern}" was normalized to diamond.`)
    const rawShredPattern = effect === 'shred' ? effectNode.getAttribute('pattern') : null
    const shredPattern = effect === 'shred' ? (rawShredPattern === 'rectangle' ? 'rectangle' : 'strip') : undefined
    if (effect === 'shred' && rawShredPattern && !['strip', 'rectangle'].includes(rawShredPattern)) warnings.add(`The shred transition pattern "${rawShredPattern}" was normalized to strip.`)
    const rawHasBounce = effect === 'flythrough' ? (effectNode.getAttribute('hasBounce')?.toLowerCase() ?? null) : null
    const hasBounce = effect === 'flythrough' ? rawHasBounce === '1' || rawHasBounce === 'true' : undefined
    if (effect === 'flythrough' && rawHasBounce !== null && !['0', '1', 'false', 'true'].includes(rawHasBounce)) warnings.add('The flythrough hasBounce value is invalid; false was used.')
    // p14 doors/window use CT_OrientationTransition; the Microsoft samples serialize the orientation in dir.
    const orientation = effect === 'split' ? (effectNode?.getAttribute('orient') as 'horz' | 'vert') || 'vert' : ['blinds', 'checker', 'comb', 'doors', 'window'].includes(effect) ? (rawDirection === 'vert' ? 'vert' : 'horz') : undefined
    if (['blinds', 'checker', 'comb', 'doors', 'window'].includes(effect) && rawDirection && !['horz', 'vert'].includes(rawDirection)) warnings.add(`The ${effect} transition orientation "${rawDirection}" was normalized to horz.`)
    const prismBoolean = (attribute: 'isInverted' | 'isContent') => {
        const value = effect === 'prism' ? effectNode.getAttribute(attribute)?.toLowerCase() : null
        if (value === null || value === undefined) return effect === 'prism' ? false : undefined
        if (['1', 'true'].includes(value)) return true
        if (['0', 'false'].includes(value)) return false
        warnings.add(`The prism ${attribute} value is invalid; false was used.`)
        return false
    }
    const isInverted = prismBoolean('isInverted')
    const isContent = prismBoolean('isContent')
    const presetBoolean = (attribute: 'invX' | 'invY'): boolean | undefined => {
        if (effect !== 'preset') return undefined
        const value = effectNode.getAttribute(attribute)?.toLowerCase()
        if (value === null || value === undefined) return false
        if (['1', 'true'].includes(value)) return true
        if (['0', 'false'].includes(value)) return false
        warnings.add(`The p15:prstTrans ${attribute} value is invalid; false was used.`)
        return false
    }
    const invertX = presetBoolean('invX')
    const invertY = presetBoolean('invY')
    const rawMorphOption = effect === 'morph' ? effectNode.getAttribute('option') : null
    const morphOption = effect === 'morph' ? ((['byObject', 'byWord', 'byChar'].includes(rawMorphOption || '') ? rawMorphOption : 'byObject') as 'byObject' | 'byWord' | 'byChar') : undefined
    if (effect === 'morph' && rawMorphOption !== 'byObject' && rawMorphOption !== 'byWord' && rawMorphOption !== 'byChar') {
        warnings.add('The morph option is missing or invalid; byObject was used.')
    }
    if (effect === 'morph' && morphOption !== 'byObject') warnings.add(`Morph ${morphOption} matching is not rendered yet; the slide uses a crossfade.`)
    // The renderer rotates the full slide plane; Office's isContent=true separates background and content.
    if (isContent) warnings.add('The prism isContent flag requests separate background and content planes; this player rotates them together.')
    return {
        effect,
        morphOption,
        presetName,
        invertX,
        invertY,
        direction,
        orientation,
        pattern,
        shredPattern,
        hasBounce,
        barOrientation: effect === 'randomBar' ? (effectNode?.getAttribute('dir') === 'vert' ? 'vertical' : 'horizontal') : undefined,
        spokes,
        isInverted,
        isContent,
        throughBlack,
        // ponytail: cap explicit transitions at one minute; raise this if real Office decks demonstrate longer playback needs.
        // Default speed timings were read from PowerPoint's SlideShowTransition.Speed and Duration properties.
        durationMs: explicitDuration !== undefined && Number.isFinite(explicitDuration) && explicitDuration >= 0 ? Math.min(explicitDuration, 60_000) : speed === 'fast' ? 500 : speed === 'slow' ? 1000 : 750
    }
}

function evaluateMotionFormula(formula: string, variables: Record<string, number>): number | undefined {
    const tokens: string[] = []
    const tokenizer = /#[A-Za-z][A-Za-z0-9_.]*|\$|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[A-Za-z][A-Za-z0-9_]*|[()+\-*/%^,]/gy
    let scan = 0
    while (scan < formula.length) {
        if (/\s/.test(formula[scan] || '')) {
            scan++
            continue
        }
        tokenizer.lastIndex = scan
        const match = tokenizer.exec(formula)
        if (!match) return undefined
        tokens.push(match[0])
        if (tokens.length > 512) return undefined
        scan = tokenizer.lastIndex
    }

    const functions: Record<string, (...args: number[]) => number> = {
        abs: Math.abs,
        acos: Math.acos,
        asin: Math.asin,
        atan: Math.atan,
        ceil: Math.ceil,
        cos: Math.cos,
        cosh: Math.cosh,
        deg: value => (value * 180) / Math.PI,
        exp: Math.exp,
        floor: Math.floor,
        ln: Math.log,
        max: Math.max,
        min: Math.min,
        rad: value => (value * Math.PI) / 180,
        sin: Math.sin,
        sinh: Math.sinh,
        sqrt: Math.sqrt,
        tan: Math.tan,
        tanh: Math.tanh
    }
    let cursor = 0
    const finite = (value: number | undefined): value is number => value !== undefined && Number.isFinite(value)
    const formulaValue = (): number | undefined => {
        let value = term()
        while (value !== undefined && (tokens[cursor] === '+' || tokens[cursor] === '-')) {
            const operator = tokens[cursor++]
            const right = term()
            if (right === undefined) return undefined
            value = operator === '+' ? value + right : value - right
            if (!Number.isFinite(value)) return undefined
        }
        return value
    }
    const factor = (depth: number): number | undefined => {
        if (depth > 255) return undefined
        const token = tokens[cursor++]
        if (token === undefined) return undefined
        if (token === '+' || token === '-') {
            const value = factor(depth + 1)
            return value === undefined ? undefined : token === '-' ? -value : value
        }
        if (token === '(') {
            const value = formulaValue()
            if (!finite(value) || tokens[cursor++] !== ')') return undefined
            return value
        }
        if (token.startsWith('#')) return variables[token.slice(1)]
        if (token === '$') return variables.$
        if (token === 'pi') return Math.PI
        if (token === 'e') return Math.E
        const numeric = Number(token)
        if (Number.isFinite(numeric)) return numeric
        const fn = functions[token]
        if (!fn || tokens[cursor++] !== '(') return undefined
        const first = formulaValue()
        if (!finite(first)) return undefined
        let args = [first]
        if (token === 'max' || token === 'min') {
            if (tokens[cursor++] !== ',') return undefined
            const second = formulaValue()
            if (!finite(second)) return undefined
            args.push(second)
        }
        if (tokens[cursor++] !== ')') return undefined
        const result = fn(...args)
        return Number.isFinite(result) ? result : undefined
    }
    const power = (): number | undefined => {
        let value = factor(0)
        while (value !== undefined && tokens[cursor] === '^') {
            cursor++
            const right = factor(0)
            if (right === undefined) return undefined
            value = value ** right
            if (!Number.isFinite(value)) return undefined
        }
        return value
    }
    const term = (): number | undefined => {
        let value = power()
        while (value !== undefined && (tokens[cursor] === '*' || tokens[cursor] === '/' || tokens[cursor] === '%')) {
            const operator = tokens[cursor++]
            const right = power()
            if (right === undefined) return undefined
            value = operator === '*' ? value * right : operator === '/' ? value / right : value % right
            if (!Number.isFinite(value)) return undefined
        }
        return value
    }

    const value = formulaValue()
    return finite(value) && cursor === tokens.length ? value : undefined
}

function parseAnimations(slideRoot: Element, theme: Record<string, string>, warnings: Set<string>, elements: PptxElement[], slideWidth: number, slideHeight: number): Pick<PptxSlide, 'automaticAnimations' | 'animationSteps' | 'animationSequence' | 'triggeredAnimations'> {
    const timing = firstDescendant(slideRoot, 'timing')
    if (!timing) return {}
    // OOXML timing node kinds and delay semantics:
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timenodevalues
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition
    const milliseconds = (value: string | null | undefined, fallback = 0) => {
        if (!value || value === 'indefinite') return fallback
        const parsed = Number(value)
        return Number.isFinite(parsed) && parsed >= 0 ? Math.min(parsed, 30_000) : fallback
    }
    const animationTiming = (timeNode: Element | undefined, effect: string): Pick<PptxAnimation, 'fillMode' | 'repeatCount' | 'repeatDurationMs' | 'autoReverse' | 'acceleration' | 'deceleration' | 'playbackSpeed' | 'reversePlayback'> | undefined => {
        const fill = timeNode?.getAttribute('fill') || 'hold'
        const autoRev = timeNode?.getAttribute('autoRev')
        if (!['hold', 'freeze', 'remove', 'transition'].includes(fill) || (autoRev && !['0', '1', 'false', 'true'].includes(autoRev))) {
            warnings.add(`${effect} animations with unsupported fill or auto-reverse timing are not rendered yet.`)
            return undefined
        }
        const countValue = timeNode?.getAttribute('repeatCount')
        const durationValue = timeNode?.getAttribute('repeatDur')
        const count = countValue === null || countValue === undefined ? (durationValue === null || durationValue === undefined ? undefined : Infinity) : countValue === 'indefinite' ? Infinity : Number(countValue) / 1000
        const rawSpeed = timeNode?.getAttribute('spd')
        const speed = rawSpeed === null || rawSpeed === undefined ? 1 : Number(rawSpeed) / 100_000
        const playbackSpeed = Number.isFinite(speed) && speed !== 0 ? speed : 1
        if (playbackSpeed !== speed) warnings.add(`${effect} animations with zero or invalid speed are played at normal speed.`)
        const playbackRate = Math.abs(playbackSpeed)
        const repeatDurationMs = durationValue === null || durationValue === undefined ? undefined : durationValue === 'indefinite' ? Infinity : Number(durationValue) / playbackRate
        if ((count !== undefined && !(count > 0)) || (repeatDurationMs !== undefined && (Number.isNaN(repeatDurationMs) || repeatDurationMs < 0))) {
            warnings.add(`${effect} animations with invalid repeat timing are not rendered yet.`)
            return undefined
        }
        const fixedPercentage = (value: string | null): number | undefined => {
            if (value === null) return 0
            const raw = value.trim()
            const percent = /^(?:\d+\.?\d*|\.\d+)%$/.test(raw) ? Number(raw.slice(0, -1)) / 100 : /^\d+$/.test(raw) ? Number(raw) / 100_000 : Number.NaN
            return Number.isFinite(percent) && percent >= 0 && percent <= 1 ? percent : undefined
        }
        const acceleration = fixedPercentage(timeNode?.getAttribute('accel') ?? null)
        const deceleration = fixedPercentage(timeNode?.getAttribute('decel') ?? null)
        if (acceleration === undefined || deceleration === undefined || acceleration + deceleration > 1) {
            warnings.add(`${effect} animations with invalid acceleration/deceleration timing are not rendered yet.`)
            return undefined
        }
        // OOXML repeatCount uses thousandths of an iteration; autoRev doubles each simple-duration cycle.
        return {
            fillMode: fill === 'remove' || fill === 'transition' ? 'remove' : (fill as 'hold' | 'freeze'),
            repeatCount: count,
            repeatDurationMs,
            autoReverse: autoRev === '1' || autoRev === 'true',
            acceleration,
            deceleration,
            playbackSpeed,
            reversePlayback: playbackSpeed < 0
        }
    }
    const animationDurationMs = (timeNode: Element | undefined, timing: Pick<PptxAnimation, 'playbackSpeed'>, fallback = 500) => {
        const duration = milliseconds(timeNode?.getAttribute('dur'), fallback) / Math.abs(timing.playbackSpeed || 1)
        if (duration > 30_000) {
            warnings.add('An animation slowed beyond 30 seconds was capped at 30 seconds.')
            return 30_000
        }
        return duration
    }
    const playbackDuration = (action: PptxAnimation) =>
        action.timingWarp
            ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs)
            : (action.effect === 'media' ? (action.mediaDurationMs ?? 0) : action.durationMs === 0 ? 0 : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (action.repeatCount ?? 1), action.repeatDurationMs ?? Infinity)) +
              (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0)
    const startDelay = (timeNode: Element | undefined) => milliseconds(child(child(timeNode, 'stCondLst'), 'cond')?.getAttribute('delay'))
    const handledEffects = new Set<Element>()
    const handledScales = new Set<Element>()
    const handledRotations = new Set<Element>()
    const handledColors = new Set<Element>()
    const handledTextAnimations = new Set<Element>()
    const handledMotions = new Set<Element>()
    const handledMediaActions = new Set<Element>()
    // OOXML also defines withGroup/afterGroup as group-level relatives of withEffect/afterEffect.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timenodevalues?view=openxml-3.0.1
    const supportedGroupTypes = new Set(['clickEffect', 'withEffect', 'afterEffect', 'withGroup', 'afterGroup'])
    const followsPreviousGroup = (type: string | null | undefined) => type === 'afterEffect' || type === 'afterGroup'
    // presetClass is a p:cTn attribute; discover nested preset nodes independently of a fixed timing-tree path.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode.presetclass?view=openxml-3.0.1
    const hasPresetClass = (timeNode: Element) => Boolean(timeNode.getAttribute('presetClass')?.trim())
    const groupType = (timeNode: Element) => timeNode.getAttribute('nodeType') || (hasPresetClass(timeNode) ? 'clickEffect' : '')
    const sceneTree = firstDescendant(slideRoot, 'spTree')
    const renderedElementIds = new Set(elements.map(element => element.id))
    const groupTargets = new Map<string, string[]>()
    const sceneChildren = (node: Element) => children(node).filter(item => ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(item.localName))
    const identityContainers: Record<string, string> = {grpSp: 'nvGrpSpPr', sp: 'nvSpPr', pic: 'nvPicPr', graphicFrame: 'nvGraphicFramePr', cxnSp: 'nvCxnSpPr'}
    const sceneNodeId = (node: Element) => {
        const identityContainer = identityContainers[node.localName]
        if (!identityContainer) return undefined
        return child(child(node, identityContainer), 'cNvPr')?.getAttribute('id')
    }
    const groupChildren = (groupId: string): string[] | undefined => {
        if (!sceneTree) return undefined
        if (groupTargets.has(groupId)) return groupTargets.get(groupId)
        let targetGroup: Element | undefined
        const findGroup = (node: Element): boolean => {
            if (node.localName === 'grpSp' && sceneNodeId(node) === groupId) {
                targetGroup = node
                return true
            }
            return sceneChildren(node).some(findGroup)
        }
        findGroup(sceneTree)
        if (!targetGroup) return undefined
        const ids: string[] = []
        const collect = (node: Element) => {
            for (const item of sceneChildren(node)) {
                if (item.localName === 'grpSp') collect(item)
                else {
                    const id = sceneNodeId(item)
                    if (id && renderedElementIds.has(id)) ids.push(id)
                }
            }
        }
        collect(targetGroup)
        groupTargets.set(groupId, ids)
        return ids
    }
    const parseGroup = (group: Element, delayMs: number, owner = child(group, 'cTn'), sequenceKey = 'main', sequenceOrder = 0, mediaWaitForEndKeys: string[] = [], mediaWaitDelayMs = 0): PptxAnimation[] => {
        const actions = new Map<string, PptxAnimation>()
        let groupTiming: (TimingCurve & {durationMs: number; playbackSpeed: number; playbackDurationMs: number; reversePlayback: boolean; autoReverse: boolean}) | undefined
        // Parent accel/decel/spd/autoRev and repeats shape the group clock; child offsets stay local to one cycle.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.childtimenodelist?view=openxml-2.19.0
        if (owner && ['accel', 'decel', 'spd', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => owner.hasAttribute(attribute))) {
            const timing = animationTiming(owner, 'Animation group')
            if (timing) {
                const groupSpeed = Math.abs(timing.playbackSpeed ?? 1)
                const rawDuration = owner?.getAttribute('dur')
                const durationMs = rawDuration === null || rawDuration === undefined ? Number.NaN : Number(rawDuration)
                if (!Number.isFinite(durationMs) || durationMs <= 0) {
                    warnings.add('A time-container timing curve or repeat without a finite positive duration is not applied.')
                } else {
                    const cycleDurationMs = (durationMs * (timing.autoReverse ? 2 : 1)) / groupSpeed
                    const repeatCount = timing.repeatCount ?? (owner.hasAttribute('repeatDur') ? Infinity : 1)
                    const playbackDurationMs = Math.min(cycleDurationMs * repeatCount, timing.repeatDurationMs ?? Infinity)
                    if (Number.isNaN(playbackDurationMs) || playbackDurationMs <= 0) {
                        warnings.add('An indefinite or empty parent time-container repeat is not applied yet.')
                    } else if (durationMs > 30_000 || (Number.isFinite(playbackDurationMs) && playbackDurationMs > 30_000)) {
                        warnings.add('A time-container timing curve or repeat beyond 30 seconds is not applied yet.')
                    } else {
                        groupTiming = {
                            durationMs,
                            playbackSpeed: groupSpeed,
                            playbackDurationMs,
                            reversePlayback: timing.reversePlayback ?? false,
                            autoReverse: timing.autoReverse ?? false,
                            acceleration: timing.acceleration ?? 0,
                            deceleration: timing.deceleration ?? 0
                        }
                    }
                }
            }
        }
        const iterate = child(owner, 'iterate')
        const belongsToOwner = (node: Element) => {
            let parent = node.parentElement
            while (parent && parent !== group.parentElement) {
                if (parent.localName === 'cTn' && (supportedGroupTypes.has(parent.getAttribute('nodeType') || '') || hasPresetClass(parent))) return parent === owner
                parent = parent.parentElement
            }
            return false
        }
        const ownedDescendants = (name: string) => descendants(group, name).filter(belongsToOwner)
        // p:charRg uses an end-exclusive range; PowerPoint requires end > st and end <= text length.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d0b4598c-f4e1-47ff-91b0-8fb5abb708f3
        // p:txEl/p:pRg is the OOXML paragraph index range:
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.textelement.paragraphindexrange
        const target = (node: Element) => {
            const shape = firstDescendant(node, 'spTgt')
            const textElement = firstDescendant(shape, 'txEl')
            const paragraphNode = firstDescendant(textElement, 'pRg')
            const characterNode = firstDescendant(textElement, 'charRg')
            const start = numberAttr(characterNode, 'st')
            const end = numberAttr(characterNode, 'end')
            const entireText = start === 0xffff_ffff && end === 0xffff_ffff
            return {
                targetId: shape?.getAttribute('spid') || '',
                paragraphRange: paragraphNode ? {start: numberAttr(paragraphNode, 'st'), end: numberAttr(paragraphNode, 'end')} : undefined,
                characterRange: characterNode ? {start: entireText ? 0 : start, end: entireText ? Number.MAX_SAFE_INTEGER : end} : undefined,
                hasUnsupportedTextRange: Boolean(textElement && ((!paragraphNode && !characterNode) || (paragraphNode && characterNode)))
            }
        }
        const hasValidCharacterRange = (animationTarget: ReturnType<typeof target>, element: PptxElement | undefined) => {
            if (!animationTarget.characterRange) return true
            const range = animationTarget.characterRange
            const textLength = element?.paragraphs.reduce((length, paragraph, index) => length + paragraph.runs.reduce((sum, run) => sum + run.text.length, 0) + (index ? 1 : 0), 0) || 0
            return textLength > 0 && Number.isInteger(range.start) && Number.isInteger(range.end) && range.start >= 0 && range.end > range.start && range.start < textLength && (range.end === Number.MAX_SAFE_INTEGER || range.end <= textLength)
        }
        const hasValidTextRange = (animationTarget: ReturnType<typeof target>, element: PptxElement | undefined) => {
            const paragraphRange = animationTarget.paragraphRange
            return (
                (!paragraphRange || Boolean(element?.paragraphs.length && Number.isInteger(paragraphRange.start) && Number.isInteger(paragraphRange.end) && paragraphRange.start >= 0 && paragraphRange.end >= paragraphRange.start && paragraphRange.end < element.paragraphs.length)) &&
                hasValidCharacterRange(animationTarget, element) &&
                (!(paragraphRange || animationTarget.characterRange) || Boolean(element?.paragraphs.length))
            )
        }
        // p:iterate applies child timing across letters, words, or shapes; tmPct is a percentage interval and tmAbs is milliseconds.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.iterate?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.iterate.type?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timepercentage?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timeabsolute?view=openxml-3.0.1
        const iterationTiming = (action: PptxAnimation, elementIteration = false) => {
            if (!iterate) return undefined
            const intervalNodes = children(iterate).filter(node => ['tmAbs', 'tmPct'].includes(node.localName))
            const intervalNode = intervalNodes[0]
            const rawInterval = intervalNode?.getAttribute('val')
            const intervalValue = rawInterval !== null && rawInterval !== undefined && /^\d+(?:\.\d+)?$/.test(rawInterval) ? Number(rawInterval) : NaN
            const intervalMs = intervalNode?.localName === 'tmAbs' ? (intervalValue <= 30_000 ? intervalValue / Math.abs(action.playbackSpeed || 1) : NaN) : intervalNode?.localName === 'tmPct' ? (action.durationMs * intervalValue) / 100_000 : NaN
            const backwardsValue = iterate.getAttribute('backwards')
            const backwards = backwardsValue === '1' || backwardsValue === 'true'
            if (intervalNodes.length !== 1 || !Number.isFinite(intervalValue) || intervalValue < 0 || (intervalNode?.localName === 'tmPct' && intervalValue > 100_000) || !Number.isFinite(intervalMs) || (elementIteration ? intervalMs < 0 : intervalMs <= 0)) {
                warnings.add(elementIteration ? 'A shape iteration with an unsupported interval was rendered without staggering.' : 'A text iteration with unsupported type or interval was rendered as a single animation.')
                return undefined
            }
            if (backwardsValue !== null && !['0', 'false', '1', 'true'].includes(backwardsValue)) {
                warnings.add(`${elementIteration ? 'A shape' : 'A text'} iteration with an invalid backwards value was rendered in forward order.`)
                return undefined
            }
            return {intervalMs, backwards}
        }
        const iterationRanges = (action: PptxAnimation): PptxAnimation['iteration'] => {
            if (!iterate) return undefined
            const type = iterate.getAttribute('type')
            if (type === 'el') return undefined
            const interval = iterationTiming(action)
            const element = elements.find(item => item.id === action.targetId)
            if (!['lt', 'wd'].includes(type || '') || !interval || !element?.paragraphs.some(paragraph => paragraph.runs.length)) {
                if (!['lt', 'wd'].includes(type || '')) warnings.add('A text iteration with an unsupported type was rendered as a single animation.')
                else if (!element?.paragraphs.some(paragraph => paragraph.runs.length)) warnings.add('A text iteration target has no text and was rendered as a single animation.')
                return undefined
            }
            const {intervalMs, backwards} = interval
            if (!['appear', 'fade', 'fontSize'].includes(action.effect) && !(action.effect === 'color' && action.colorProperty === 'style.color')) {
                warnings.add(`Text iteration for ${action.effect} animation is not rendered yet.`)
                return undefined
            }
            if (!Intl.Segmenter) {
                warnings.add('Text iteration is unavailable because Intl.Segmenter is not supported by this browser.')
                return undefined
            }
            const segmenter = new Intl.Segmenter(undefined, {granularity: type === 'lt' ? 'grapheme' : 'word'})
            const ranges: Array<{start: number; end: number}> = []
            let overflow = false
            let paragraphOffset = 0
            for (const [paragraphIndex, paragraph] of element.paragraphs.entries()) {
                const text = paragraph.runs.map(run => run.text).join('')
                const paragraphSelected = !action.paragraphRange || (paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
                if (paragraphSelected) {
                    for (const segment of segmenter.segment(text)) {
                        if (type === 'wd' && !segment.isWordLike) continue
                        const start = paragraphOffset + segment.index
                        const end = start + segment.segment.length
                        const selection = action.characterRange
                        if (!selection || (start < selection.end && end > selection.start)) ranges.push({start, end})
                        if (ranges.length > 10_000) {
                            overflow = true
                            break
                        }
                    }
                }
                paragraphOffset += text.length + (paragraphIndex < element.paragraphs.length - 1 ? 1 : 0)
                if (overflow) break
            }
            if (overflow) {
                warnings.add('Text iteration exceeding 10,000 units was rendered as a single animation.')
                return undefined
            }
            return ranges.length > 1 ? {intervalMs, ranges, backwards} : undefined
        }
        const elementIterationActions = (action: PptxAnimation): PptxAnimation[] | undefined => {
            if (iterate?.getAttribute('type') !== 'el') return undefined
            const targetIds = groupChildren(action.targetId) || (renderedElementIds.has(action.targetId) ? [action.targetId] : [])
            if (!targetIds.length) {
                warnings.add('A shape iteration target has no rendered shapes and was skipped.')
                return []
            }
            if (targetIds.length === 1) return [{...action, targetId: targetIds[0]!}]
            const interval = iterationTiming(action, true)
            return targetIds.map((targetId, index) => ({
                ...action,
                targetId,
                delayMs: action.delayMs + (interval?.intervalMs || 0) * (interval?.backwards ? targetIds.length - index - 1 : index)
            }))
        }
        const addMediaAction = (node: Element, targetId: string, mediaCommand: NonNullable<PptxAnimation['mediaCommand']>, options: Pick<PptxAnimation, 'mediaStartSeconds' | 'mediaVolume' | 'mediaMuted' | 'mediaDurationMs' | 'mediaSlideCount' | 'mediaWaitForEnd'> = {}) => {
            const element = elements.find(item => item.id === targetId)
            if (!element || !['audio', 'video'].includes(element.kind)) {
                warnings.add('A media timeline action without an embedded audio or video target was skipped.')
                handledMediaActions.add(node)
                return
            }
            const timeNode = firstDescendant(node, 'cTn')
            const conditions = children(child(timeNode, 'stCondLst'), 'cond')
            const condition = conditions[0]
            const delay = condition?.getAttribute('delay')
            if (conditions.length > 1 || (condition && (Array.from(condition.attributes).some(attribute => attribute.localName !== 'delay') || condition.children.length || delay === 'indefinite' || (delay !== null && (!Number.isFinite(Number(delay)) || Number(delay) < 0))))) {
                warnings.add('A media timeline action with event-based or invalid start conditions was skipped.')
                handledMediaActions.add(node)
                return
            }
            const key = `media:${timeNode?.getAttribute('id') || actions.size}`
            actions.set(key, {
                targetId,
                effect: 'media',
                direction: 'in',
                mediaCommand,
                ...options,
                durationMs: 0,
                delayMs: delayMs + milliseconds(delay)
            })
            handledMediaActions.add(node)
        }
        for (const mediaNode of [...ownedDescendants('audio'), ...ownedDescendants('video')].filter(node => node.namespaceURI === PRESENTATION_NS)) {
            const media = child(mediaNode, 'cMediaNode')
            if (!media) {
                warnings.add('A timed media playback action without a common media node was skipped.')
                handledMediaActions.add(mediaNode)
                continue
            }
            const animationTarget = target(media)
            if (!animationTarget.targetId || animationTarget.hasUnsupportedTextRange || animationTarget.paragraphRange || animationTarget.characterRange) {
                warnings.add('A timed media playback action with an unsupported target was skipped.')
                handledMediaActions.add(mediaNode)
                continue
            }
            const volumeValue = media.getAttribute('vol')
            const volume = volumeValue === null ? undefined : Number(volumeValue) / 100_000
            const muteValue = media.getAttribute('mute')?.toLowerCase()
            const muted = muteValue === undefined ? undefined : ['1', 'true'].includes(muteValue)
            const slideCountValue = media.getAttribute('numSld')
            const parsedSlideCount = slideCountValue === null ? undefined : Number(slideCountValue)
            const mediaSlideCount = parsedSlideCount !== undefined && Number.isInteger(parsedSlideCount) && parsedSlideCount > 0 && parsedSlideCount <= 0xffff_ffff ? parsedSlideCount : undefined
            if (volumeValue !== null && (!Number.isFinite(volume) || volume! < 0 || volume! > 1)) warnings.add('An invalid media timeline volume was ignored.')
            if (muteValue !== undefined && !['0', '1', 'false', 'true'].includes(muteValue)) warnings.add('An invalid media timeline mute value was ignored.')
            if (slideCountValue !== null && mediaSlideCount === undefined) warnings.add('An invalid media timeline numSld value was ignored.')
            if (mediaSlideCount && mediaSlideCount > 1 && elements.find(element => element.id === animationTarget.targetId)?.kind === 'video') warnings.add('Cross-slide video playback is not supported; the video stops when leaving its slide.')
            const mediaTimeNode = firstDescendant(media, 'cTn')
            const mediaDurationValue = mediaTimeNode?.getAttribute('dur')
            const endConditions = children(child(mediaTimeNode, 'endCondLst'), 'cond')
            const hasSupportedEndCondition =
                elements.find(element => element.id === animationTarget.targetId)?.kind === 'audio' &&
                endConditions.every(condition => {
                    const delay = condition.getAttribute('delay')
                    return condition.getAttribute('evt') === 'onStopAudio' && Boolean(firstDescendant(condition, 'sldTgt')) && (delay === null || (/^\d+$/.test(delay) && Number(delay) === 0))
                })
            const indefiniteDuration = mediaDurationValue === 'indefinite'
            const parsedMediaDuration = mediaDurationValue === null || mediaDurationValue === undefined ? 0 : /^\d+$/.test(mediaDurationValue) ? Number(mediaDurationValue) : undefined
            const mediaDurationMs = parsedMediaDuration !== undefined && Number.isFinite(parsedMediaDuration) && parsedMediaDuration <= 2_147_483_625 ? parsedMediaDuration : undefined
            if (parsedMediaDuration !== undefined && mediaDurationMs === undefined) warnings.add('A media timeline duration beyond the supported ST_TLTime range was skipped.')
            if (mediaDurationValue !== null && mediaDurationValue !== undefined && !indefiniteDuration && parsedMediaDuration === undefined) warnings.add('An invalid media timeline duration was skipped.')
            if (endConditions.length && !hasSupportedEndCondition) warnings.add('A media timeline end condition other than slide-targeted onStopAudio is not synchronized yet.')
            addMediaAction(mediaNode, animationTarget.targetId, 'play', {
                mediaVolume: volume !== undefined && volume >= 0 && volume <= 1 ? volume : undefined,
                mediaMuted: muteValue === undefined || ['0', '1', 'false', 'true'].includes(muteValue) ? muted : undefined,
                mediaDurationMs,
                mediaSlideCount,
                mediaWaitForEnd: indefiniteDuration && (!endConditions.length || hasSupportedEndCondition)
            })
        }
        for (const command of ownedDescendants('cmd').filter(node => node.namespaceURI === PRESENTATION_NS)) {
            const animationTarget = target(command)
            const commandType = command.getAttribute('type')
            const source = command.getAttribute('cmd')?.trim() || ''
            const playFrom = /^PlayFrom\(\s*(\d+(?:\.\d+)?)\s*\)$/i.exec(source)
            const normalizedCommand = source.toLowerCase()
            const mediaCommand = playFrom ? 'play' : normalizedCommand === 'togglepause' ? 'togglePause' : normalizedCommand === 'pause' ? 'pause' : normalizedCommand === 'stop' ? 'stop' : undefined
            if (!mediaCommand || commandType !== 'call' || !animationTarget.targetId || animationTarget.hasUnsupportedTextRange || animationTarget.paragraphRange || animationTarget.characterRange) {
                warnings.add('An unsupported media or object command in the slide timeline was skipped.')
                handledMediaActions.add(command)
                continue
            }
            const startSeconds = playFrom ? Number(playFrom[1]) : undefined
            if (startSeconds !== undefined && (!Number.isFinite(startSeconds) || startSeconds > 86_400)) {
                warnings.add('A PlayFrom media command with an out-of-range start time was skipped.')
                handledMediaActions.add(command)
                continue
            }
            addMediaAction(command, animationTarget.targetId, mediaCommand, {mediaStartSeconds: startSeconds})
        }
        for (const effect of ownedDescendants('animEffect')) {
            const animationTarget = target(effect)
            const textElement = elements.find(item => item.id === animationTarget.targetId)
            const transition = effect.getAttribute('transition')
            const filter = effect.getAttribute('filter')
            const wipe = /^wipe\((right|left|up|down)\)$/.exec(filter || '')
            // Microsoft Open XML uses filter="blinds(horizontal)" for an object entrance effect.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animateeffect?view=openxml-3.0.1
            const blinds = /^blinds\((horizontal|vertical)\)$/.exec(filter || '')
            // MS-OE376 documents checkerboard(across/down) as object-effect filters.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const checker = /^checkerboard\((across|down)\)$/.exec(filter || '')
            // MS-OE376 lists random horizontal and vertical object bars.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const randomBar = /^randombar\((horizontal|vertical)\)$/.exec(filter || '')
            // MS-OE376 lists four diagonal strips variants as object filters.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const strips = /^strips\((downLeft|upLeft|downRight|upRight)\)$/.exec(filter || '')
            // MS-OE376 lists the four horizontal/vertical barn object filters.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const barn = /^barn\((in|out)(Horizontal|Vertical)\)$/.exec(filter || '')
            // Office lists these four filters; translate their object across the slide as an entrance/exit.
            // The from-side direction is interpreted from the named OOXML filter.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const slide = /^slide\(from(Top|Bottom|Left|Right)\)$/.exec(filter || '')
            // MS-OE376 lists dissolve as an image-filter animation.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const dissolve = filter === 'dissolve'
            // MS-OE376 names these object filters and their subtypes.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const shapeFilter = /^(circle|diamond|box|plus)\((in|out)\)$/.exec(filter || '')
            // MS-OE376 defines wedge and wheel(1/2/3/4/8) object filters.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a96dab70-2e72-4319-928d-0eb4b275ce58
            const wheel = /^wheel\(([1-4]|8)\)$/.exec(filter || '')
            const wedge = filter === 'wedge'
            if (!animationTarget.targetId) {
                warnings.add('An object animation has no shape target and was skipped.')
                continue
            }
            if (animationTarget.hasUnsupportedTextRange) {
                warnings.add('This text range animation target is malformed or unsupported.')
                continue
            }
            if (!hasValidCharacterRange(animationTarget, textElement)) {
                warnings.add('A character-range animation exceeds the target text length and was skipped.')
                continue
            }
            if (['in', 'out'].includes(transition || '') && (['fade', 'none', 'cut'].includes(filter || '') || wipe || blinds || checker || randomBar || strips || barn || slide || dissolve || shapeFilter || wheel || wedge)) {
                const paragraphWipe = Boolean(wipe && animationTarget.paragraphRange && !animationTarget.characterRange)
                const characterWipe = Boolean(wipe && animationTarget.characterRange && !animationTarget.paragraphRange)
                // Microsoft documents pRg with checkerboard(across) as a valid paragraph-targeted animation.
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.paragraphindexrange?view=openxml-3.0.1
                const unsupportedTextMask = randomBar || strips || barn || slide || dissolve || shapeFilter || wheel || wedge || (wipe && !paragraphWipe && !characterWipe)
                if (unsupportedTextMask && (animationTarget.paragraphRange || animationTarget.characterRange)) {
                    warnings.add('Text-range slide, random bars, strips, barn, dissolve, wheel, and shape-mask animations are not rendered yet.')
                    handledEffects.add(effect)
                    continue
                }
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.progress
                const timeNode = firstDescendant(effect, 'cTn')
                const timing = animationTiming(timeNode, 'Object')
                if (!timing) {
                    handledEffects.add(effect)
                    continue
                }
                const key = `${animationTarget.targetId}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: wheel || wedge ? 'wheel' : slide ? 'slide' : strips ? 'strips' : barn ? 'barn' : dissolve ? 'dissolve' : shapeFilter ? 'shape' : randomBar ? 'randomBars' : checker ? 'checker' : blinds ? 'blinds' : wipe ? 'wipe' : filter === 'fade' ? 'fade' : 'appear',
                    direction: transition as 'in' | 'out',
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing,
                    wipeDirection: wipe?.[1] as PptxAnimation['wipeDirection'],
                    blindsOrientation: blinds?.[1] as PptxAnimation['blindsOrientation'],
                    checkerOrientation: checker?.[1] === 'across' ? 'horz' : checker ? 'vert' : undefined,
                    randomBarOrientation: randomBar?.[1] as PptxAnimation['randomBarOrientation'],
                    stripsDirection: strips ? ({downLeft: 'ld', upLeft: 'lu', downRight: 'rd', upRight: 'ru'} as const)[strips[1] as 'downLeft' | 'upLeft' | 'downRight' | 'upRight'] : undefined,
                    barnOrientation: barn?.[2]?.toLowerCase() as PptxAnimation['barnOrientation'],
                    barnMotion: barn?.[1] as PptxAnimation['barnMotion'],
                    shapeFilter: shapeFilter?.[1] as PptxAnimation['shapeFilter'],
                    shapeDirection: shapeFilter?.[2] as PptxAnimation['shapeDirection'],
                    slideFrom: slide?.[1]?.toLowerCase() as PptxAnimation['slideFrom'],
                    wheelSpokes: wheel ? (Number(wheel[1]) as PptxAnimation['wheelSpokes']) : wedge ? 1 : undefined
                })
                handledEffects.add(effect)
            } else {
                warnings.add(`The ${transition || 'unknown'} ${filter || 'unknown'} object animation is not rendered.`)
            }
        }
        for (const animation of ownedDescendants('animScale')) {
            const animationTarget = target(animation)
            const timeNode = firstDescendant(animation, 'cTn')
            const element = elements.find(item => item.id === animationTarget.targetId)
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                warnings.add('A scale animation with an unsupported target or text range was skipped.')
                continue
            }
            const timing = animationTiming(timeNode, 'Scale')
            if (!timing) {
                handledScales.add(animation)
                continue
            }
            // DrawingML scale coordinates are in 100,000ths; the SDK example uses by=150000 for a 150% scale.
            // https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.presentation.animatescale?view=openxml-3.0.1
            const pair = (node: Element): [number, number] => [numberAttr(node, 'x', 100_000) / 100_000, numberAttr(node, 'y', 100_000) / 100_000]
            const fromNode = child(animation, 'from')
            const toNode = child(animation, 'to')
            const byNode = child(animation, 'by')
            const scaleFrom = fromNode ? pair(fromNode) : ([1, 1] as [number, number])
            const by = byNode ? pair(byNode) : undefined
            const scaleTo = toNode ? pair(toNode) : by ? ([scaleFrom[0] * by[0], scaleFrom[1] * by[1]] as [number, number]) : undefined
            if (!scaleTo || [...scaleFrom, ...scaleTo].some(value => !Number.isFinite(value) || value <= 0 || value > 10)) {
                warnings.add('A scale animation with missing or out-of-range values was skipped.')
                continue
            }
            const key = `${animationTarget.targetId}:scale:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'scale',
                direction: 'in',
                scaleFrom,
                scaleTo,
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + startDelay(timeNode),
                ...timing
            })
            handledScales.add(animation)
        }
        for (const animation of ownedDescendants('animRot')) {
            const behavior = child(animation, 'cBhvr')
            const animationTarget = target(behavior || animation)
            const element = elements.find(item => item.id === animationTarget.targetId)
            const names = descendants(child(behavior, 'attrNameLst'), 'attrName')
            const property = names[0]?.textContent?.trim()
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element) || !['r', 'ppt_r', 'style.rotation'].includes(property || '')) {
                warnings.add('A rotation animation with an unsupported target or property was skipped.')
                continue
            }
            const rawFrom = animation.getAttribute('from')
            const rawTo = animation.getAttribute('to')
            const rawBy = animation.getAttribute('by')
            const parseAngle = (value: string | null) => (value !== null && /^[-+]?\d+$/.test(value) && Number.isFinite(Number(value)) && Math.abs(Number(value)) <= 2_147_483_647 ? Number(value) / 60_000 : undefined)
            const from = rawFrom === null ? undefined : parseAngle(rawFrom)
            const to = rawTo === null ? undefined : parseAngle(rawTo)
            const by = rawBy === null ? undefined : parseAngle(rawBy)
            const valid = (from !== undefined && to !== undefined && rawBy === null) || (from !== undefined && by !== undefined && rawTo === null) || (rawFrom === null && to !== undefined && rawBy === null) || (rawFrom === null && rawTo === null && by !== undefined)
            if (!valid) {
                warnings.add('A rotation animation with an invalid from/to/by combination was skipped.')
                continue
            }
            const rotationFrom = from ?? 0
            const rotationTo = to ?? rotationFrom + (by ?? 0)
            const timeNode = child(behavior, 'cTn') || firstDescendant(animation, 'cTn')
            const timing = animationTiming(timeNode, 'Rotation')
            if (!timing) {
                handledRotations.add(animation)
                continue
            }
            const key = `${animationTarget.targetId}:rotation:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'rotation',
                direction: 'in',
                rotationFrom,
                rotationTo,
                rotationRelative: from === undefined && to === undefined,
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + startDelay(timeNode),
                ...timing
            })
            handledRotations.add(animation)
        }
        for (const animation of ownedDescendants('animClr')) {
            const behavior = child(animation, 'cBhvr')
            const animationTarget = target(behavior || animation)
            const property = descendants(child(behavior, 'attrNameLst'), 'attrName')[0]?.textContent?.trim()
            const element = elements.find(item => item.id === animationTarget.targetId)
            const colorSpace = animation.getAttribute('clrSpc') || 'rgb'
            const colorDirection = animation.getAttribute('dir') || 'cw'
            const invalidTarget =
                !element ||
                animationTarget.hasUnsupportedTextRange ||
                ((animationTarget.paragraphRange || animationTarget.characterRange) && property !== 'style.color') ||
                !hasValidCharacterRange(animationTarget, element) ||
                (property === 'fillcolor' && (element?.kind !== 'shape' || /gradient\(/i.test(element.fill))) ||
                (property === 'stroke.color' && (!['shape', 'line'].includes(element?.kind || '') || /gradient\(/i.test(element.stroke))) ||
                (property === 'shadow.color' && !element?.shadow && !element?.paragraphs.some(paragraph => paragraph.runs.some(run => run.shadow))) ||
                (property === 'style.color' && (!element?.paragraphs.length || (element.kind === 'table' && animationTarget.paragraphRange)))
            if (!animationTarget.targetId || invalidTarget || !['fillcolor', 'style.color', 'stroke.color', 'shadow.color'].includes(property || '')) {
                warnings.add('A color animation with an unsupported target or property was skipped.')
                handledColors.add(animation)
                continue
            }
            if (!['rgb', 'hsl'].includes(colorSpace) || !['cw', 'ccw'].includes(colorDirection)) {
                warnings.add('A color animation with an unsupported color space or hue direction was skipped.')
                handledColors.add(animation)
                continue
            }
            const colorValue = (wrapper: Element | undefined) => animationColorChannels(children(wrapper)[0], theme, colorSpace as 'rgb' | 'hsl')
            const fromNode = child(animation, 'from')
            const toNode = child(animation, 'to')
            const byNode = child(animation, 'by')
            const black: [number, number, number, number] = [0, 0, 0, 1]
            const colorFrom = fromNode ? colorValue(fromNode) : black
            const explicitTo = toNode ? colorValue(toNode) : undefined
            const colorBy = byNode ? colorValue(byNode) : undefined
            if (!colorFrom || (toNode && !explicitTo) || (!toNode && byNode && !colorBy) || (!toNode && !byNode)) {
                warnings.add('A color animation with missing or invalid from/to/by colors was skipped.')
                handledColors.add(animation)
                continue
            }
            const colorTo =
                explicitTo ||
                (colorFrom.map((channel, index) => {
                    const sum = channel + colorBy![index]!
                    return index === 0 && colorSpace === 'hsl' ? sum : Math.max(0, Math.min(index === 3 ? 1 : colorSpace === 'rgb' && index < 3 ? 255 : 1, sum))
                }) as [number, number, number, number])
            const timeNode = child(behavior, 'cTn') || firstDescendant(animation, 'cTn')
            const timing = animationTiming(timeNode, 'Color')
            if (!timing) {
                handledColors.add(animation)
                continue
            }
            const key = `${animationTarget.targetId}:color:${property}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'color',
                direction: 'in',
                colorProperty: property as PptxAnimation['colorProperty'],
                colorFrom,
                colorTo,
                colorSpace: colorSpace as PptxAnimation['colorSpace'],
                colorDirection: colorDirection as PptxAnimation['colorDirection'],
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + startDelay(timeNode),
                ...timing
            })
            handledColors.add(animation)
        }
        for (const animation of ownedDescendants('anim')) {
            // Microsoft lists style.fontWeight among p:tav formula-capable target attributes.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa
            // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
            // p:tav elements are time/value keypoints; numeric opacity supports the discrete and linear calculation modes.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timeanimatevalue?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animatebehaviorcalculatemodevalues?view=openxml-3.0.1
            const behavior = child(animation, 'cBhvr')
            const animationTarget = target(behavior || animation)
            const property = descendants(child(behavior, 'attrNameLst'), 'attrName')[0]?.textContent?.trim()
            const element = elements.find(item => item.id === animationTarget.targetId)
            const timeNode = child(behavior, 'cTn') || firstDescendant(animation, 'cTn')
            const tavList = child(animation, 'tavLst')
            const calculationMode = animation.getAttribute('calcmode') || 'lin'
            if (
                !animationTarget.targetId ||
                !['style.fontSize', 'style.fontWeight', 'style.opacity', 'fill.opacity', 'stroke.opacity', 'shadow.opacity'].includes(property || '') ||
                !element ||
                (['style.fontSize', 'style.fontWeight'].includes(property || '') && !element.paragraphs.length) ||
                (property === 'fill.opacity' && (element.kind !== 'shape' || element.fill.includes('gradient('))) ||
                (property === 'stroke.opacity' && (!['shape', 'line'].includes(element.kind) || element.stroke.includes('gradient('))) ||
                (property === 'shadow.opacity' && !element.shadow && !element.paragraphs.some(paragraph => paragraph.runs.some(run => run.shadow))) ||
                animationTarget.hasUnsupportedTextRange ||
                !hasValidCharacterRange(animationTarget, element) ||
                (animation.getAttribute('valueType') && animation.getAttribute('valueType') !== 'num') ||
                !['lin', 'discrete'].includes(calculationMode) ||
                (calculationMode === 'discrete' && !tavList) ||
                (tavList && children(tavList).some(node => node.localName !== 'tav'))
            ) {
                warnings.add('A generic object animation with an unsupported target or interpolation was skipped.')
                handledTextAnimations.add(animation)
                continue
            }
            const rawFrom = animation.getAttribute('from')
            const rawTo = animation.getAttribute('to')
            const rawBy = animation.getAttribute('by')
            const parseFixedPercentage = (raw: string | null): number | undefined => {
                if (raw === null) return undefined
                const value = raw.trim()
                // ST_PositiveFixedPercentage accepts a trailing percent or 1000ths of a percent; Office writes the latter.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c1f1feac-e34c-48d5-b2e2-f67bb67113e7
                if (/^\d+(?:\.\d+)?%$/.test(value)) return Number(value.slice(0, -1)) / 100
                if (/^\d+$/.test(value)) return Number(value) / 100_000
                return undefined
            }
            const keyframeNodes = tavList ? children(tavList) : []
            // ponytail: cap at 256 keypoints to bound parser and per-frame work; raise only when real decks need denser curves.
            const parsedKeyframes =
                tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes
                          .map(node => {
                              const valueNode = children(child(node, 'val'))[0]
                              const offset = parseFixedPercentage(node.getAttribute('tm'))
                              const rawValue = valueNode && ['fltVal', 'strVal'].includes(valueNode.localName) ? valueNode.getAttribute('val') : null
                              const value = rawValue?.trim() ? Number(rawValue) : Number.NaN
                              const formula = node.hasAttribute('fmla') ? node.getAttribute('fmla')?.trim() : undefined
                              return offset !== undefined && offset >= 0 && offset <= 1 && Number.isFinite(value) && value >= 0 && value <= 1000
                                  && (formula === undefined || Boolean(formula && formula.length <= 16_384))
                                  ? {offset, value, formula}
                                  : undefined
                          })
                          .filter((frame): frame is {offset: number; value: number; formula: string | undefined} => Boolean(frame))
                    : undefined
            const validKeyframeTrack = Boolean(parsedKeyframes?.length && parsedKeyframes.length === keyframeNodes.length && parsedKeyframes.every((frame, index) => index === 0 || frame.offset > parsedKeyframes[index - 1]!.offset))
            const keyframes = validKeyframeTrack && parsedKeyframes?.every(frame => frame.formula === undefined)
                ? parsedKeyframes.map(({offset, value}) => ({offset, value}))
                : undefined
            const validKeyframes = Boolean(keyframes?.length)
            if (!['style.fontSize', 'style.fontWeight'].includes(property || '')) {
                const hasTextRange = Boolean(animationTarget.paragraphRange || animationTarget.characterRange)
                if (hasTextRange && property !== 'style.opacity') {
                    warnings.add('Text-range fill, stroke, and shadow opacity animations are not rendered yet.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const validOpacityKeyframes = Boolean(validKeyframes && keyframes?.every(frame => frame.value <= 1))
                const validValues = tavList ? validOpacityKeyframes && rawFrom === null && rawTo === null && rawBy === null : !(rawTo !== null && rawBy !== null) && (rawTo !== null || rawBy !== null) && !(rawFrom !== null && rawTo === null && rawBy === null)
                const from = rawFrom === null ? 0 : Number(rawFrom)
                const to = rawTo !== null ? Number(rawTo) : rawBy !== null ? from + Number(rawBy) : undefined
                if (!validValues || (!tavList && (to === undefined || !Number.isFinite(from) || !Number.isFinite(to) || from < 0 || from > 1 || to < 0 || to > 1))) {
                    warnings.add('A generic opacity animation with missing or out-of-range values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Opacity')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const rangeKey = hasTextRange ? ':' + (animationTarget.paragraphRange?.start ?? '') + ':' + (animationTarget.paragraphRange?.end ?? '') + ':' + (animationTarget.characterRange?.start ?? '') + ':' + (animationTarget.characterRange?.end ?? '') : ''
                actions.set(animationTarget.targetId + ':opacity:' + property + rangeKey, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'opacity',
                    opacityProperty: property as PptxAnimation['opacityProperty'],
                    direction: 'in',
                    opacityFrom: keyframes?.[0]?.value ?? from,
                    opacityTo: keyframes?.at(-1)?.value ?? to,
                    opacityKeyframes: keyframes,
                    opacityKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['opacityKeyframeMode']>) : undefined,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (property === 'style.fontWeight') {
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                const validFontWeightKeyframes = Boolean(validKeyframeTrack && parsedKeyframes?.every(frame => frame.value >= 1 && frame.value <= 1000))
                const validValues = tavList
                    ? validFontWeightKeyframes && rawFrom === null && rawTo === null && rawBy === null && (!hasFormulaKeyframes || calculationMode === 'lin')
                    : rawFrom !== null && rawTo !== null && rawBy === null
                const from = parsedKeyframes?.[0]?.value ?? (rawFrom === null ? Number.NaN : Number(rawFrom))
                const to = parsedKeyframes?.at(-1)?.value ?? (rawTo === null ? Number.NaN : Number(rawTo))
                if (!validValues || !Number.isFinite(from) || !Number.isFinite(to) || from < 1 || from > 1000 || to < 1 || to > 1000) {
                    warnings.add(hasFormulaKeyframes
                        ? 'A generic text font-weight formula animation with unsupported interpolation or invalid numeric keyframes was skipped.'
                        : 'A generic text font-weight animation with missing or out-of-range numeric values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Text font weight')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                let fontWeightFormulaSamples: PptxAnimation['fontWeightFormulaSamples']
                if (hasFormulaKeyframes && parsedKeyframes) {
                    const sampleCount = Math.min(900, Math.max(60, Math.ceil(durationMs / 1000 * 60)))
                    fontWeightFormulaSamples = []
                    for (let index = 0; index <= sampleCount; index++) {
                        const progress = index / sampleCount
                        const rightIndex = parsedKeyframes.findIndex(frame => frame.offset >= progress)
                        const left = rightIndex < 0
                            ? parsedKeyframes.at(-1)!
                            : parsedKeyframes[Math.max(0, rightIndex - (parsedKeyframes[rightIndex]!.offset === progress ? 0 : 1))]!
                        const right = rightIndex < 0 ? left : parsedKeyframes[rightIndex]!
                        const formula = parsedKeyframes.filter(frame => frame.offset <= progress && frame.formula !== undefined).at(-1)?.formula
                        const ratio = right.offset === left.offset ? 0 : (progress - left.offset) / (right.offset - left.offset)
                        const baseValue = left.value + (right.value - left.value) * Math.max(0, Math.min(1, ratio))
                        const value = formula
                            ? evaluateMotionFormula(formula, {
                                  ppt_x: element.x / slideWidth,
                                  ppt_y: element.y / slideHeight,
                                  ppt_w: element.width / slideWidth,
                                  ppt_h: element.height / slideHeight,
                                  'style.fontWeight': baseValue,
                                  $: progress
                              })
                            : baseValue
                        if (value === undefined || value < 1 || value > 1000) {
                            fontWeightFormulaSamples = undefined
                            break
                        }
                        fontWeightFormulaSamples.push({offset: progress, value})
                    }
                    if (!fontWeightFormulaSamples) {
                        warnings.add('A generic text font-weight formula used an unsupported variable or produced an out-of-range value; animation was skipped.')
                        handledTextAnimations.add(animation)
                        continue
                    }
                }
                const key = `${animationTarget.targetId}:fontWeight:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'fontWeight',
                    direction: 'in',
                    fontWeightFrom: from,
                    fontWeightTo: to,
                    fontWeightKeyframes: keyframes,
                    fontWeightFormulaSamples,
                    fontWeightKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['fontWeightKeyframeMode']>) : undefined,
                    durationMs,
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            const validFontSizeKeyframes = Boolean(validKeyframes && keyframes?.every(frame => frame.value > 0))
            const validValues = tavList ? validFontSizeKeyframes && rawFrom === null && rawTo === null && rawBy === null : !(rawTo !== null && rawBy !== null) && (rawTo !== null || rawBy !== null) && !(rawFrom !== null && rawTo === null && rawBy === null)
            const from = keyframes?.[0]?.value ?? (rawFrom === null ? 1 : Number(rawFrom))
            const to = keyframes?.at(-1)?.value ?? (rawTo === null ? (rawBy === null ? undefined : from + Number(rawBy)) : Number(rawTo))
            if (!validValues || to === undefined || !Number.isFinite(from) || !Number.isFinite(to) || from <= 0 || to <= 0 || from > 100 || to > 100) {
                warnings.add('A generic text animation with missing or out-of-range font sizes was skipped.')
                handledTextAnimations.add(animation)
                continue
            }
            const timing = animationTiming(timeNode, 'Text font size')
            if (!timing) {
                handledTextAnimations.add(animation)
                continue
            }
            const key = `${animationTarget.targetId}:fontSize:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'fontSize',
                direction: 'in',
                fontSizeFrom: from,
                fontSizeTo: to,
                fontSizeKeyframes: keyframes,
                fontSizeKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['fontSizeKeyframeMode']>) : undefined,
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + startDelay(timeNode),
                ...timing
            })
            handledTextAnimations.add(animation)
        }
        for (const animation of ownedDescendants('animMotion')) {
            const animationTarget = target(animation)
            const timeNode = firstDescendant(animation, 'cTn')
            const element = elements.find(item => item.id === animationTarget.targetId)
            const path = animation.getAttribute('path')?.trim() || ''
            const origin = animation.getAttribute('origin') || 'parent'
            const fromNode = child(animation, 'from')
            const toNode = child(animation, 'to')
            const byNode = child(animation, 'by')
            const hasPosition = Boolean(fromNode || toNode || byNode)
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                warnings.add('A motion animation with an unsupported target or text range was skipped.')
                continue
            }
            if (!path && !hasPosition && ['layout', 'parent'].includes(origin)) {
                // An empty p:animMotion path has no visual effect but still consumes its behavior duration.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
                const timing = animationTiming(timeNode, 'Motion')
                if (timing)
                    actions.set(`${animationTarget.targetId}:timing:${timeNode?.getAttribute('id') || ''}`, {
                        targetId: animationTarget.targetId,
                        effect: 'timingOnly',
                        direction: 'in',
                        durationMs: animationDurationMs(timeNode, timing),
                        delayMs: delayMs + startDelay(timeNode),
                        ...timing
                    })
                handledMotions.add(animation)
                continue
            }
            if (!['layout', 'parent'].includes(origin) || (path ? hasPosition : !hasPosition)) {
                warnings.add('Motion animations without a supported layout or parent path are not rendered yet.')
                continue
            }
            const timing = animationTiming(timeNode, 'Motion')
            if (!timing) {
                handledMotions.add(animation)
                continue
            }
            const angleSource = animation.getAttribute('rAng')
            const angleText = angleSource?.trim()
            const rotationAngle = angleText === undefined ? 0 : /^[+-]?\d+$/.test(angleText) ? Number(angleText) : Number.NaN
            const readCoordinate = (raw: string | null): number | undefined => {
                if (raw === null) return undefined
                const value = raw.trim()
                // MS-OI29500 ST_Percentage is either a percent value or integer thousandths of a percent.
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.rotationcenter?view=openxml-3.0.1
                if (/^[+-]?(?:\d+\.?\d*|\.\d+)%$/.test(value)) return Number(value.slice(0, -1)) / 100
                if (/^[+-]?\d+$/.test(value)) return Number(value) / 100_000
                return undefined
            }
            const readPosition = (node: Element | undefined): [number, number] | undefined => {
                if (!node) return undefined
                const x = readCoordinate(node.getAttribute('x'))
                const y = readCoordinate(node.getAttribute('y'))
                return x !== undefined && y !== undefined && Number.isFinite(x) && Number.isFinite(y) ? [x, y] : undefined
            }
            const rotationCenter = readPosition(child(animation, 'rCtr'))
            // PowerPoint defaults rAng to zero; nonzero values need an explicit slide-space rCtr.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            if (!Number.isInteger(rotationAngle) || rotationAngle < -2_147_483_648 || rotationAngle > 2_147_483_647 || (rotationAngle !== 0 && (origin !== 'layout' || !rotationCenter))) {
                warnings.add('A motion animation with an invalid rotation angle or unsupported rotation center was skipped.')
                continue
            }
            const angleRadians = (rotationAngle * Math.PI) / (60_000 * 180)
            const angleCosine = Math.cos(angleRadians),
                angleSine = Math.sin(angleRadians)
            const rotateMotionOffset = (x: number, y: number): [number, number] => {
                if (rotationAngle === 0 || !rotationCenter || !element) return [x, y]
                const centerX = rotationCenter[0] * slideWidth,
                    centerY = rotationCenter[1] * slideHeight
                const dx = element.x + x - centerX,
                    dy = element.y + y - centerY
                return [centerX + dx * angleCosine - dy * angleSine - element.x, centerY + dx * angleSine + dy * angleCosine - element.y]
            }
            if (!path) {
                const hasFrom = Boolean(fromNode),
                    hasTo = Boolean(toNode),
                    hasBy = Boolean(byNode)
                const validCombination = hasTo !== hasBy && (hasFrom || hasTo || hasBy)
                const from = readPosition(fromNode),
                    to = readPosition(toNode),
                    by = readPosition(byNode)
                if (!element || !validCombination || (fromNode && !from) || (toNode && !to) || (byNode && !by)) {
                    warnings.add('A motion animation with an invalid from/to/by position was skipped.')
                    continue
                }
                const start: [number, number] = from || (origin === 'layout' ? [0, 0] : [element.x / slideWidth, element.y / slideHeight])
                const end: [number, number] = to || [start[0] + (by?.[0] || 0), start[1] + (by?.[1] || 0)]
                const toOffset = (coordinates: [number, number]): [number, number] => (origin === 'layout' ? [coordinates[0] * slideWidth, coordinates[1] * slideHeight] : [coordinates[0] * slideWidth - element.x, coordinates[1] * slideHeight - element.y])
                const [startX, startY] = toOffset(start)
                const [endX, endY] = toOffset(end)
                const [rotatedStartX, rotatedStartY] = rotateMotionOffset(startX, startY)
                const [rotatedEndX, rotatedEndY] = rotateMotionOffset(endX, endY)
                const limit = 100 * Math.max(slideWidth, slideHeight)
                const distance = Math.hypot(rotatedEndX - rotatedStartX, rotatedEndY - rotatedStartY)
                if ([rotatedStartX, rotatedStartY, rotatedEndX, rotatedEndY].some(value => !Number.isFinite(value) || Math.abs(value) > limit) || !Number.isFinite(distance)) {
                    warnings.add('A motion animation with out-of-range from/to/by coordinates was skipped.')
                    continue
                }
                const key = `${animationTarget.targetId}:motion:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'motion',
                    direction: 'in',
                    motionPath: [
                        {x: rotatedStartX, y: rotatedStartY, distance: 0},
                        {x: rotatedEndX, y: rotatedEndY, distance}
                    ],
                    motionPathLength: distance,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing
                })
                handledMotions.add(animation)
                continue
            }
            if (path.length > 100_000) {
                warnings.add('A motion path longer than 100,000 characters was skipped.')
                continue
            }
            const pathEditMode = animation.getAttribute('pathEditMode') || 'relative'
            if (!['relative', 'fixed'].includes(pathEditMode)) {
                warnings.add('A motion path with an unsupported pathEditMode was skipped.')
                continue
            }
            // Office motion path commands use slide-size fractions and absolute/relative coordinates.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            const tokens: Array<string | {coordinates: [number, number]} | {formulas: [string, string]}> = []
            const numberPattern = /[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/iy
            let scan = 0
            let invalidPathSyntax = false
            let tooManyTokens = false
            let unsupportedFormula = false
            let dynamicFormula = false
            const formulaVariables: Record<string, number> = element
                ? {
                      ppt_x: element.x / slideWidth,
                      ppt_y: element.y / slideHeight,
                      ppt_w: element.width / slideWidth,
                      ppt_h: element.height / slideHeight
                  }
                : {ppt_x: Number.NaN, ppt_y: Number.NaN, ppt_w: Number.NaN, ppt_h: Number.NaN}
            while (scan < path.length && !invalidPathSyntax) {
                const character = path[scan]
                if (character && /[\s,]/.test(character)) {
                    scan++
                    continue
                }
                if (character && /[MLCZE]/i.test(character)) {
                    const action = character
                    tokens.push(action)
                    scan++
                    if (action.toLowerCase() === 'e') break
                } else if (character === '(') {
                    const start = ++scan
                    let depth = 1
                    while (scan < path.length && depth > 0) {
                        if (path[scan] === '(') depth++
                        else if (path[scan] === ')') depth--
                        scan++
                        if (depth > 256) {
                            invalidPathSyntax = true
                            break
                        }
                    }
                    if (depth !== 0 || invalidPathSyntax) {
                        invalidPathSyntax = true
                        break
                    }
                    const formulaPair = path.slice(start, scan - 1)
                    let formulaDepth = 0
                    const commaSeparators: number[] = []
                    const spaceSeparators: number[] = []
                    for (let index = 0; index < formulaPair.length; index++) {
                        if (formulaPair[index] === '(') formulaDepth++
                        else if (formulaPair[index] === ')') formulaDepth--
                        else if (formulaDepth === 0 && formulaPair[index] === ',') commaSeparators.push(index)
                        else if (formulaDepth === 0 && /\s/.test(formulaPair[index] || '') && (index === 0 || !/\s/.test(formulaPair[index - 1] || ''))) spaceSeparators.push(index)
                    }
                    const dynamicCoordinates = formulaPair.includes('$')
                    if (dynamicCoordinates) dynamicFormula = true
                    if (formulaPair.length > 16_384 || commaSeparators.length > 1) {
                        unsupportedFormula = true
                        invalidPathSyntax = true
                        break
                    }
                    let coordinates: [number, number] | undefined
                    let formulas: [string, string] | undefined
                    const separators = commaSeparators.length ? commaSeparators : spaceSeparators
                    for (const separator of separators) {
                        const xFormula = formulaPair.slice(0, separator).trim()
                        const yFormula = formulaPair.slice(separator + 1).trim()
                        const variables = dynamicCoordinates ? {...formulaVariables, $: 0.5} : formulaVariables
                        const x = evaluateMotionFormula(xFormula, variables)
                        const y = evaluateMotionFormula(yFormula, variables)
                        if (x !== undefined && y !== undefined) {
                            coordinates = [x, y]
                            if (dynamicCoordinates) formulas = [xFormula, yFormula]
                            break
                        }
                    }
                    if (!coordinates) {
                        unsupportedFormula = true
                        invalidPathSyntax = true
                        break
                    }
                    tokens.push(formulas ? {formulas} : {coordinates})
                } else {
                    numberPattern.lastIndex = scan
                    const number = numberPattern.exec(path)
                    if (!number || !number[0]) {
                        invalidPathSyntax = true
                        break
                    }
                    tokens.push(number[0])
                    scan = numberPattern.lastIndex
                }
                if (tokens.length > 1024) {
                    tooManyTokens = true
                    break
                }
            }
            if (tooManyTokens) {
                warnings.add('A motion path with more than 1,024 coordinates was skipped.')
                continue
            }
            if (invalidPathSyntax || !tokens.length || !element) {
                if (unsupportedFormula) warnings.add('A motion path with an unsupported or invalid coordinate formula was skipped.')
                else warnings.add('An invalid or unsupported motion path was skipped.')
                continue
            }
            if (!Number.isFinite(formulaVariables.ppt_x) || !Number.isFinite(formulaVariables.ppt_y) || !Number.isFinite(formulaVariables.ppt_w) || !Number.isFinite(formulaVariables.ppt_h)) {
                warnings.add('An invalid or unsupported motion path was skipped.')
                continue
            }
            // Relative layout paths travel with the target; fixed paths stay anchored to the slide.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            const point = (x: number, y: number): [number, number] => {
                if (origin === 'parent' && element.motionParent) {
                    const [a, b, c, d] = element.motionParent.matrix
                    const localX = x * slideWidth - element.motionParent.x
                    const localY = y * slideHeight - element.motionParent.y
                    return [a * localX + c * localY, b * localX + d * localY]
                }
                return origin === 'layout' ? (pathEditMode === 'fixed' ? [x * slideWidth - element.x, y * slideHeight - element.y] : [x * slideWidth, y * slideHeight]) : [x * slideWidth - element.x, y * slideHeight - element.y]
            }
            const limit = 100 * Math.max(slideWidth, slideHeight)
            const buildPath = (progress: number) => {
                const pathPoints: NonNullable<PptxAnimation['motionPath']> = []
                const pathControls: [number, number][] = []
                let cursor = 0
                let current: [number, number] | undefined
                let pathStart: [number, number] | undefined
                let invalid = false
                let drawingSegments = 0
                let distance = 0
                const append = (coordinates: [number, number], snap = false) => {
                    const [x, y] = point(...coordinates)
                    const previous = pathPoints.at(-1)
                    if (!Number.isFinite(x) || !Number.isFinite(y) || Math.abs(x) > limit || Math.abs(y) > limit) {
                        invalid = true
                        return
                    }
                    if (previous && !snap) distance += Math.hypot(x - previous.x, y - previous.y)
                    pathPoints.push({x, y, distance})
                }
                const readPoint = (): [number, number] | undefined => {
                    const x = tokens[cursor++]
                    if (x === undefined) return undefined
                    if (typeof x !== 'string') {
                        if ('coordinates' in x) return x.coordinates
                        const variables = {...formulaVariables, $: progress}
                        const formulaX = evaluateMotionFormula(x.formulas[0], variables)
                        const formulaY = evaluateMotionFormula(x.formulas[1], variables)
                        return formulaX === undefined || formulaY === undefined ? undefined : [formulaX, formulaY]
                    }
                    const y = tokens[cursor++]
                    if (typeof y !== 'string' || !Number.isFinite(Number(x)) || !Number.isFinite(Number(y))) return undefined
                    return [Number(x), Number(y)]
                }
                while (cursor < tokens.length && !invalid) {
                    const action = tokens[cursor++]
                    if (typeof action !== 'string') {
                        invalid = true
                        break
                    }
                    if (action.toLowerCase() === 'e') break
                    if (action.toLowerCase() === 'm') {
                        const coordinates = readPoint()
                        if (!coordinates) {
                            invalid = true
                            break
                        }
                        current = action === 'm' ? [coordinates[0], coordinates[1]] : coordinates
                        pathStart ??= [...current]
                        // M snaps to a point; disconnected subpaths add no travel distance.
                        append(current, true)
                        continue
                    }
                    if (!current) {
                        invalid = true
                        break
                    }
                    if (action.toLowerCase() === 'l') {
                        const coordinates = readPoint()
                        if (!coordinates) {
                            invalid = true
                            break
                        }
                        const end = action === 'l' ? ([current[0] + coordinates[0], current[1] + coordinates[1]] as [number, number]) : coordinates
                        append(end)
                        current = end
                        drawingSegments++
                    } else if (action.toLowerCase() === 'c') {
                        const values = [readPoint(), readPoint(), readPoint()]
                        if (values.some(value => !value)) {
                            invalid = true
                            break
                        }
                        const [control1Input, control2Input, endInput] = values as [[number, number], [number, number], [number, number]]
                        const relative = action === 'c'
                        const control1 = relative ? ([current[0] + control1Input[0], current[1] + control1Input[1]] as [number, number]) : control1Input
                        const control2 = relative ? ([current[0] + control2Input[0], current[1] + control2Input[1]] as [number, number]) : control2Input
                        const end = relative ? ([current[0] + endInput[0], current[1] + endInput[1]] as [number, number]) : endInput
                        pathControls.push(control1, control2, end)
                        const from = point(...current),
                            first = point(...control1),
                            second = point(...control2),
                            to = point(...end)
                        // ponytail: 32 samples per cubic bound per-frame lookup size; increase or make adaptive if subpixel pacing error appears in calibrated decks.
                        for (let step = 1; step <= 32; step++) {
                            const t = step / 32,
                                inverse = 1 - t
                            const x = inverse ** 3 * from[0] + 3 * inverse ** 2 * t * first[0] + 3 * inverse * t ** 2 * second[0] + t ** 3 * to[0]
                            const y = inverse ** 3 * from[1] + 3 * inverse ** 2 * t * first[1] + 3 * inverse * t ** 2 * second[1] + t ** 3 * to[1]
                            const previous = pathPoints.at(-1)
                            if (previous) distance += Math.hypot(x - previous.x, y - previous.y)
                            pathPoints.push({x, y, distance})
                        }
                        current = end
                        drawingSegments++
                    } else if (action.toLowerCase() === 'z') {
                        if (!pathStart) {
                            invalid = true
                            break
                        }
                        append(pathStart)
                        current = [...pathStart]
                        drawingSegments++
                    } else {
                        invalid = true
                        break
                    }
                    if (drawingSegments > 128) invalid = true
                }
                if (
                    !invalid &&
                    pathControls.some(coordinates => {
                        const [x, y] = rotateMotionOffset(...point(...coordinates))
                        return !Number.isFinite(x) || !Number.isFinite(y) || Math.abs(x) > limit || Math.abs(y) > limit
                    })
                )
                    invalid = true
                if (!invalid)
                    for (const pathPoint of pathPoints) {
                        const [x, y] = rotateMotionOffset(pathPoint.x, pathPoint.y)
                        pathPoint.x = x
                        pathPoint.y = y
                        if (!Number.isFinite(x) || !Number.isFinite(y) || Math.abs(x) > limit || Math.abs(y) > limit) invalid = true
                    }
                return {points: pathPoints, length: distance, invalid: invalid || !pathPoints.length || !Number.isFinite(distance)}
            }
            const pathPosition = (path: NonNullable<PptxAnimation['motionPath']>, length: number, progress: number): [number, number] => {
                if (path.length === 1 || length <= 0) return [path.at(-1)!.x, path.at(-1)!.y]
                const distance = length * progress
                let low = 0,
                    high = path.length
                while (low < high) {
                    const middle = (low + high) >>> 1
                    if (path[middle]!.distance <= distance) low = middle + 1
                    else high = middle
                }
                const from = path[low - 1],
                    to = path[low]
                if (!from) return [path[0]!.x, path[0]!.y]
                if (!to) return [from.x, from.y]
                const span = to.distance - from.distance
                const ratio = span ? (distance - from.distance) / span : 0
                return [from.x + (to.x - from.x) * ratio, from.y + (to.y - from.y) * ratio]
            }
            const durationMs = animationDurationMs(timeNode, timing)
            let pathGeometry = buildPath(0)
            let motionPathSamples: PptxAnimation['motionPathSamples']
            if (dynamicFormula) {
                const desiredSamples = Math.max(60, Math.ceil((durationMs / 1000) * 60))
                const sampleCount = Math.min(900, desiredSamples)
                if (sampleCount < desiredSamples) warnings.add('A long animation-progress motion path is sampled at a reduced rate.')
                motionPathSamples = []
                for (let index = 0; index <= sampleCount; index++) {
                    const progress = index / sampleCount
                    const geometry = index === 0 ? pathGeometry : buildPath(progress)
                    if (geometry.invalid) {
                        pathGeometry = geometry
                        break
                    }
                    const [x, y] = pathPosition(geometry.points, geometry.length, progress)
                    motionPathSamples.push({progress, x, y})
                }
                if (motionPathSamples.length !== sampleCount + 1) {
                    warnings.add('A dynamic motion-path formula produced an invalid or out-of-range coordinate; animation was skipped.')
                    continue
                }
            }
            if (pathGeometry.invalid) {
                warnings.add('A motion animation with out-of-range coordinates was skipped.')
                continue
            }
            const key = `${animationTarget.targetId}:motion:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'motion',
                direction: 'in',
                motionPath: pathGeometry.points,
                motionPathLength: pathGeometry.length,
                motionPathSamples,
                durationMs,
                delayMs: delayMs + startDelay(timeNode),
                ...timing
            })
            handledMotions.add(animation)
        }
        for (const set of ownedDescendants('set')) {
            const attribute = firstDescendant(set, 'attrName')?.textContent?.trim()
            const value = firstDescendant(set, 'strVal')?.getAttribute('val')
            const behavior = child(set, 'cBhvr')
            const animationTarget = target(behavior || set)
            const element = elements.find(item => item.id === animationTarget.targetId)
            // p:set holds one fixed style.fontSize factor; Microsoft's animation example uses 1.5 for 150%.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.setbehavior?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
            if (attribute === 'style.fontSize') {
                if (!animationTarget.targetId || !element?.paragraphs.length || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element)) {
                    warnings.add('A style.fontSize p:set animation with an unsupported target was skipped.')
                    continue
                }
                const to = child(set, 'to')
                const rawFontSize = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const fontSize = rawFontSize?.trim() ? Number(rawFontSize) : Number.NaN
                if (!Number.isFinite(fontSize) || fontSize < 0 || fontSize > 100) {
                    warnings.add('A style.fontSize p:set animation with a missing or out-of-range value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set font size')
                if (!timing) continue
                const key = `${animationTarget.targetId}:fontSize:set:${timeNode?.getAttribute('id') || actions.size}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'fontSize',
                    direction: 'in',
                    fontSizeFrom: fontSize,
                    fontSizeTo: fontSize,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (['style.opacity', 'fill.opacity', 'stroke.opacity', 'shadow.opacity'].includes(attribute || '')) {
                const property = attribute as NonNullable<PptxAnimation['opacityProperty']>
                const unsupportedTarget =
                    !animationTarget.targetId ||
                    !element ||
                    animationTarget.hasUnsupportedTextRange ||
                    ((animationTarget.paragraphRange || animationTarget.characterRange) && property !== 'style.opacity') ||
                    !hasValidCharacterRange(animationTarget, element) ||
                    (property === 'fill.opacity' && (element?.kind !== 'shape' || /gradient\(/i.test(element.fill))) ||
                    (property === 'stroke.opacity' && (!['shape', 'line'].includes(element?.kind || '') || /gradient\(/i.test(element.stroke))) ||
                    (property === 'shadow.opacity' && !element?.shadow && !element?.paragraphs.some(paragraph => paragraph.runs.some(run => run.shadow)))
                if (unsupportedTarget) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported target was skipped.`)
                    continue
                }
                const to = child(set, 'to')
                const rawValue = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const opacity = rawValue?.trim() ? Number(rawValue) : Number.NaN
                if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) {
                    warnings.add(`A ${attribute} p:set animation with a missing or out-of-range value was skipped.`)
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set opacity')
                if (!timing) continue
                const key = `${animationTarget.targetId}:opacity:${property}:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'opacity',
                    direction: 'in',
                    opacityProperty: property,
                    opacityFrom: opacity,
                    opacityTo: opacity,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (['fillcolor', 'style.color', 'stroke.color', 'shadow.color'].includes(attribute || '')) {
                const property = attribute as NonNullable<PptxAnimation['colorProperty']>
                const unsupportedTarget =
                    !animationTarget.targetId ||
                    !element ||
                    animationTarget.hasUnsupportedTextRange ||
                    animationTarget.paragraphRange ||
                    animationTarget.characterRange ||
                    !hasValidCharacterRange(animationTarget, element) ||
                    (property === 'fillcolor' && (element?.kind !== 'shape' || /gradient\(/i.test(element.fill))) ||
                    (property === 'stroke.color' && (!['shape', 'line'].includes(element?.kind || '') || /gradient\(/i.test(element.stroke))) ||
                    (property === 'shadow.color' && !element?.shadow && !element?.paragraphs.some(paragraph => paragraph.runs.some(run => run.shadow))) ||
                    (property === 'style.color' && !element?.paragraphs.length)
                if (unsupportedTarget) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported target was skipped.`)
                    continue
                }
                const colorValue = child(child(set, 'to'), 'clrVal')
                const color = animationColorChannels(children(colorValue)[0], theme, 'rgb')
                if (!color) {
                    warnings.add(`A ${attribute} p:set animation with a missing or invalid color was skipped.`)
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set color')
                if (!timing) continue
                const key = `${animationTarget.targetId}:color:${property}:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'color',
                    direction: 'in',
                    colorProperty: property,
                    colorFrom: color,
                    colorTo: color,
                    colorSpace: 'rgb',
                    colorDirection: 'cw',
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + startDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute !== 'style.visibility') {
                warnings.add(`A p:set animation for ${attribute || 'an unknown property'} is not rendered yet.`)
                continue
            }
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element)) {
                warnings.add('A style.visibility p:set animation with an unsupported target was skipped.')
                continue
            }
            if (!['visible', 'hidden'].includes(value || '')) {
                warnings.add('A style.visibility p:set animation with an invalid value was skipped.')
                continue
            }
            const key = `${animationTarget.targetId}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            if (!actions.has(key)) {
                const timeNode = firstDescendant(set, 'cTn')
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'appear',
                    direction: value === 'visible' ? 'in' : 'out',
                    durationMs: 0,
                    delayMs: delayMs + startDelay(timeNode)
                })
            }
        }
        return [...actions.values()].flatMap((action, actionIndex) => {
            const sequencedAction = {
                ...action,
                sequenceKey,
                sequenceOrder,
                mediaActionKey: action.effect === 'media' ? `${sequenceKey}:${sequenceOrder}:${actionIndex}:${action.targetId}` : undefined,
                mediaWaitForEndKeys: mediaWaitForEndKeys.length ? mediaWaitForEndKeys : undefined,
                mediaWaitDelayMs: mediaWaitForEndKeys.length ? mediaWaitDelayMs + action.delayMs - delayMs : undefined
            }
            const elementActions = elementIterationActions(sequencedAction)
            const expandedActions = elementActions ? elementActions.map(elementAction => ({...elementAction, sequenceKey, sequenceOrder})) : [{...sequencedAction, iteration: iterationRanges(sequencedAction)}]
            if (!groupTiming) return expandedActions
            return expandedActions.map(expandedAction => {
                if (expandedAction.effect === 'media') {
                    warnings.add('A time-container timing curve or repeat on timed media is not applied yet.')
                    return expandedAction
                }
                const childOffsetMs = expandedAction.delayMs - delayMs
                const childDurationMs = playbackDuration(expandedAction)
                if (!Number.isFinite(childOffsetMs) || childOffsetMs < 0 || !Number.isFinite(childDurationMs) || childOffsetMs + childDurationMs > groupTiming!.durationMs + 0.01) {
                    warnings.add('A time-container timing curve or repeat is not applied because child timing exceeds its finite duration.')
                    return expandedAction
                }
                const scale = groupTiming!.durationMs / groupTiming!.playbackSpeed
                const firstChildOffsetMs = groupTiming!.reversePlayback ? childOffsetMs + childDurationMs : childOffsetMs
                const wallProgress = inverseTimingProgress(groupTiming!, firstChildOffsetMs / groupTiming!.durationMs)
                const wallStartMs = (groupTiming!.reversePlayback ? 1 - wallProgress : wallProgress) * scale
                return {
                    ...expandedAction,
                    delayMs: delayMs + wallStartMs,
                    timingWarp: {
                        acceleration: groupTiming!.acceleration,
                        deceleration: groupTiming!.deceleration,
                        groupOffsetMs: delayMs,
                        groupDurationMs: groupTiming!.durationMs,
                        groupSpeed: groupTiming!.playbackSpeed,
                        groupReversePlayback: groupTiming!.reversePlayback,
                        groupAutoReverse: groupTiming!.autoReverse,
                        groupPlaybackDurationMs: groupTiming!.playbackDurationMs,
                        childOffsetMs,
                        groupEndDelayMs: delayMs + groupTiming!.playbackDurationMs
                    }
                }
            })
        })
    }

    const mainTimeNode = descendants(timing, 'cTn').find(node => node.getAttribute('nodeType') === 'mainSeq')
    const mainSequence = mainTimeNode?.parentElement?.localName === 'seq' ? mainTimeNode.parentElement : undefined
    const navigationCondition = (listName: 'nextCondLst' | 'prevCondLst', eventName: 'onNext' | 'onPrev') => {
        const list = child(mainSequence, listName)
        if (!list) return undefined
        const conditions = children(list, 'cond')
        const isSupported = (condition: Element) => condition.getAttribute('evt') === eventName && (!condition.hasAttribute('delay') || condition.getAttribute('delay') === '0') && Boolean(firstDescendant(condition, 'sldTgt'))
        const supported = conditions.some(isSupported)
        if (conditions.some(condition => !isSupported(condition))) warnings.add(`Main animation sequence ${listName} supports only immediate slide-target ${eventName} conditions.`)
        return supported
    }
    const nextAction = mainSequence?.getAttribute('nextAc') || 'none'
    const previousAction = mainSequence?.getAttribute('prevAc') || 'none'
    if (mainSequence && !['none', 'seek'].includes(nextAction)) warnings.add(`Main animation sequence next action "${nextAction}" is not supported.`)
    if (mainSequence && !['none', 'skipTimed'].includes(previousAction)) warnings.add(`Main animation sequence previous action "${previousAction}" is not supported.`)
    const animationSequence = mainSequence
        ? {
              nextAction: (['none', 'seek'].includes(nextAction) ? nextAction : 'none') as 'none' | 'seek',
              advancesOnNext: navigationCondition('nextCondLst', 'onNext'),
              rewindsOnPrevious: navigationCondition('prevCondLst', 'onPrev')
          }
        : undefined
    const mainChildTnList = child(mainTimeNode, 'childTnLst')
    const groups = children(mainChildTnList)
    const mainGroupFor = (node: Element) => {
        let current: Element | null = node
        while (current?.parentElement && current.parentElement !== mainChildTnList) current = current.parentElement
        return current?.parentElement === mainChildTnList ? current : undefined
    }
    const nestedInteractiveNode = (node: Element) => {
        let parent = node.parentElement
        while (parent && parent !== mainTimeNode) {
            if (parent.localName === 'cTn' && parent.getAttribute('nodeType') === 'interactiveSeq') return true
            parent = parent.parentElement
        }
        return false
    }
    const mainEvents =
        mainTimeNode && mainChildTnList
            ? descendants(mainChildTnList, 'cTn').filter(node => {
                  const directGroup = groups.some(group => child(group, 'cTn') === node)
                  return !nestedInteractiveNode(node) && (directGroup || hasPresetClass(node))
              })
            : []
    const automaticAnimations: PptxAnimation[] = []
    const steps: PptxAnimation[][] = []
    const triggeredAnimations: NonNullable<PptxSlide['triggeredAnimations']> = []
    let currentStep: PptxAnimation[] | undefined
    for (const [sequenceOrder, timeNode] of mainEvents.entries()) {
        const group = mainGroupFor(timeNode)
        if (!group) continue
        const type = groupType(timeNode)
        if (type === 'clickEffect') {
            currentStep = []
            steps.push(currentStep)
        } else if (!['withEffect', 'afterEffect', 'withGroup', 'afterGroup'].includes(type || '')) {
            continue
        }
        const sequence = currentStep || automaticAnimations
        const mediaWaitForEndKeys = [...new Set(sequence.flatMap(action => (action.mediaWaitForEnd && action.mediaActionKey ? [action.mediaActionKey] : [])))]
        const precedingEnd = followsPreviousGroup(type) ? sequence.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), 0) : 0
        let groupStartDelay = 0
        let parent: Element | null = timeNode
        while (parent && parent !== mainTimeNode) {
            if (parent.localName === 'cTn') groupStartDelay += startDelay(parent)
            parent = parent.parentElement
        }
        sequence.push(...parseGroup(group, precedingEnd + groupStartDelay, timeNode, 'main', sequenceOrder, followsPreviousGroup(type) ? mediaWaitForEndKeys : [], groupStartDelay))
    }
    const interactiveSequences = descendants(timing, 'cTn').filter(node => node.getAttribute('nodeType') === 'interactiveSeq')
    const shapeTriggerEvents = new Set(['onClick', 'onDblClick', 'onMouseOver', 'onMouseOut'])
    const slideTriggerEvents = new Set(['onNext', 'onPrev'])
    for (const [interactiveIndex, interactive] of interactiveSequences.entries()) {
        const conditions = children(child(interactive, 'stCondLst'), 'cond')
        const triggerForCondition = (condition: Element): NonNullable<PptxSlide['triggeredAnimations']>[number]['trigger'] | undefined => {
            const event = condition.getAttribute('evt') || ''
            const shapeTarget = firstDescendant(condition, 'spTgt')
            if (shapeTriggerEvents.has(event) && shapeTarget) {
                const id = shapeTarget.getAttribute('spid') || ''
                return id && elements.some(element => element.id === id) ? {type: 'shape', id, event: event as 'onClick' | 'onDblClick' | 'onMouseOver' | 'onMouseOut'} : undefined
            }
            if (slideTriggerEvents.has(event) && firstDescendant(condition, 'sldTgt')) {
                return {type: 'slide', event: event as 'onNext' | 'onPrev'}
            }
            return undefined
        }
        const triggerMatch = conditions.map(condition => ({condition, trigger: triggerForCondition(condition)})).find(item => item.trigger)
        const triggerCondition = triggerMatch?.condition
        const trigger = triggerMatch?.trigger
        if (!triggerCondition || !trigger) {
            warnings.add('An interactive animation without a supported shape or slide navigation trigger was skipped.')
            continue
        }
        if (conditions.some(condition => condition !== triggerCondition)) warnings.add('Only one trigger condition per interactive animation is supported.')

        const markers = descendants(interactive, 'cTn').filter(node => {
            if (!supportedGroupTypes.has(groupType(node)) && !hasPresetClass(node)) return false
            let parent = node.parentElement
            while (parent) {
                if (parent.localName === 'cTn' && parent.getAttribute('nodeType') === 'interactiveSeq') return parent === interactive
                parent = parent.parentElement
            }
            return false
        })
        const triggerSteps: PptxAnimation[][] = []
        let currentStep: PptxAnimation[] | undefined
        const inheritedDelay = (timeNode: Element) => {
            let delay = 0
            let parent: Element | null = timeNode
            while (parent && parent !== interactive) {
                if (parent.localName === 'cTn') delay += startDelay(parent)
                parent = parent.parentElement
            }
            return delay
        }
        for (const [sequenceOrder, timeNode] of markers.entries()) {
            const type = groupType(timeNode)
            if (type === 'clickEffect' || !currentStep) {
                currentStep = []
                triggerSteps.push(currentStep)
            }
            const group = timeNode.parentElement
            if (!group) continue
            const mediaWaitForEndKeys = [...new Set(currentStep.flatMap(action => (action.mediaWaitForEnd && action.mediaActionKey ? [action.mediaActionKey] : [])))]
            const precedingEnd = followsPreviousGroup(type) ? currentStep.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), 0) : 0
            const groupStartDelay = inheritedDelay(timeNode)
            currentStep.push(...parseGroup(group, precedingEnd + groupStartDelay, timeNode, `interactive:${interactiveIndex}`, sequenceOrder, followsPreviousGroup(type) ? mediaWaitForEndKeys : [], groupStartDelay))
        }
        if (triggerSteps.some(step => step.length)) triggeredAnimations.push({trigger, steps: triggerSteps})
    }
    if (descendants(timing, 'animEffect').some(effect => !handledEffects.has(effect))) {
        warnings.add('Interactive animation sequences are not played yet.')
    }
    if (descendants(timing, 'animScale').some(animation => !handledScales.has(animation))) {
        warnings.add('Scale animations outside supported click/automatic groups are not rendered yet.')
    }
    if (descendants(timing, 'animMotion').some(animation => !handledMotions.has(animation))) {
        warnings.add('Object motion, emphasis, and property animations are not rendered yet.')
    }
    if (descendants(timing, 'animRot').some(animation => !handledRotations.has(animation))) {
        warnings.add('Unsupported rotation animations are not rendered yet.')
    }
    if (descendants(timing, 'animClr').some(animation => !handledColors.has(animation))) {
        warnings.add('Unsupported color animations are not rendered yet.')
    }
    if (descendants(timing, 'anim').some(animation => !handledTextAnimations.has(animation))) {
        warnings.add('Unsupported generic motion, emphasis, and property animations are not rendered yet.')
    }
    if (['audio', 'video', 'cmd'].some(name => descendants(timing, name).some(node => node.namespaceURI === PRESENTATION_NS && !handledMediaActions.has(node)))) {
        warnings.add('Media playback outside supported slide timing groups is not rendered yet.')
    }
    return {
        automaticAnimations: automaticAnimations.length ? automaticAnimations : undefined,
        animationSteps: steps.length ? steps : undefined,
        animationSequence,
        triggeredAnimations: triggeredAnimations.length ? triggeredAnimations : undefined
    }
}

async function parsePicture(shape: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>): Promise<PptxElement> {
    const properties = child(shape, 'spPr')
    const transform = shapeTransform(properties)
    const blip = firstDescendant(shape, 'blip')
    const imageRelationId = namespacedAttr(blip, 'embed') || ''
    const imageRelation = rels.get(imageRelationId)
    const target = imageRelation && !imageRelation.external ? imageRelation.target : undefined
    const imageUrl = await imageUrlFor(target, zip, urls, imageUrlCache)
    if (blip && !imageUrl) warnings.add('An image reference could not be resolved.')
    const videoFile = firstDescendant(shape, 'videoFile')
    const audioFile = firstDescendant(shape, 'audioFile')
    const mediaNode = firstDescendant(shape, 'media')
    const trimNode = child(mediaNode, 'trim')
    const fadeNode = child(mediaNode, 'fade')
    const trimStartMs = trimNode ? mediaTimeOffsetMs(trimNode.getAttribute('st')) : 0
    const trimEndMs = trimNode ? mediaTimeOffsetMs(trimNode.getAttribute('end')) : 0
    const fadeInMs = fadeNode ? mediaTimeOffsetMs(fadeNode.getAttribute('in')) : 0
    const fadeOutMs = fadeNode ? mediaTimeOffsetMs(fadeNode.getAttribute('out')) : 0
    const mediaTrim = trimNode && trimStartMs !== undefined && trimEndMs !== undefined && (trimStartMs > 0 || trimEndMs > 0) ? {startMs: trimStartMs, endMs: trimEndMs} : undefined
    const mediaFade = fadeNode && fadeInMs !== undefined && fadeOutMs !== undefined && (fadeInMs > 0 || fadeOutMs > 0) ? {inMs: fadeInMs, outMs: fadeOutMs} : undefined
    if (trimNode && !mediaTrim && (trimStartMs === undefined || trimEndMs === undefined)) warnings.add('An invalid embedded media trim range was ignored; the full media is available.')
    if (fadeNode && !mediaFade && (fadeInMs === undefined || fadeOutMs === undefined)) warnings.add('An invalid embedded media fade was ignored.')
    // MS-PPTX CT_Media gives r:link precedence when p14:media has both r:link and r:embed.
    const mediaRelationId = namespacedAttr(mediaNode, 'link') || namespacedAttr(mediaNode, 'embed') || namespacedAttr(videoFile, 'link') || namespacedAttr(audioFile, 'link')
    const mediaRelation = mediaRelationId ? rels.get(mediaRelationId) : undefined
    let mediaUrl: string | undefined
    let mediaKind: 'video' | 'audio' | undefined
    let embeddedMime: string | undefined
    let mediaFormat: string | undefined
    if (videoFile || audioFile || mediaNode) {
        const mediaTarget = mediaRelation && !mediaRelation.external ? mediaRelation.target : undefined
        const mime = mediaTarget ? mediaMime(mediaTarget) : undefined
        embeddedMime = mime
        mediaFormat = mediaTarget?.split('.').at(-1)?.toLowerCase()
        mediaKind = audioFile || mime?.startsWith('audio/') ? 'audio' : 'video'
        if (mediaTarget && mime) {
            mediaUrl = imageUrlCache.get(mediaTarget)
            if (!mediaUrl) {
                const media = zip.file(mediaTarget)
                if (media) {
                    const blob = await media.async('blob')
                    mediaUrl = URL.createObjectURL(new Blob([blob], {type: mime}))
                    imageUrlCache.set(mediaTarget, mediaUrl)
                    urls.push(mediaUrl)
                }
            }
            if (!mediaUrl) warnings.add(`The embedded ${mediaKind} media part is missing; the preview image is shown when available.`)
            if (typeof document !== 'undefined' && !document.createElement(mediaKind).canPlayType(mime)) {
                warnings.add(`This browser may not support embedded ${mediaKind} format ${mediaTarget.split('.').at(-1)?.toLowerCase() || 'unknown'}.`)
            }
        } else if (mediaRelation?.external) {
            warnings.add(`Externally linked ${mediaKind} media was not loaded; only embedded media is supported.`)
        } else if (!mime) {
            warnings.add(`The embedded ${mediaKind} format is not supported; the preview image is shown when available.`)
        } else {
            warnings.add(`The embedded ${mediaKind} media reference could not be resolved; the preview image is shown when available.`)
        }
    }
    const cropNode = firstDescendant(shape, 'srcRect')
    const cropValue = (name: string) => {
        const raw = cropNode?.getAttribute(name)
        if (raw === null || raw === undefined) return 0
        const value = Number(raw)
        return Number.isFinite(value) ? value / 100_000 : Number.NaN
    }
    const imageCrop = cropNode
        ? {
              left: cropValue('l'),
              top: cropValue('t'),
              right: cropValue('r'),
              bottom: cropValue('b')
          }
        : undefined
    const validCrop = !!imageCrop && Object.values(imageCrop).every(value => value >= 0 && value < 1) && imageCrop.left + imageCrop.right < 1 && imageCrop.top + imageCrop.bottom < 1
    if (imageCrop && !validCrop) warnings.add('An invalid picture crop rectangle was ignored.')
    const nonVisual = firstDescendant(shape, 'cNvPr')
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'Picture',
        kind: mediaUrl ? mediaKind! : imageUrl ? 'picture' : 'placeholder',
        geometry: 'rect',
        ...transform,
        fill: 'transparent',
        stroke: 'transparent',
        strokeWidth: 0,
        imageUrl,
        mediaUrl,
        mediaType: mediaKind,
        mediaMime: embeddedMime,
        mediaFormat,
        mediaTrim,
        mediaFade,
        imageCrop: validCrop ? imageCrop : undefined,
        shadow: outerShadow(properties, theme),
        paragraphs: [],
        margins: {left: 0, right: 0, top: 0, bottom: 0},
        verticalAlign: 'top'
    }
}

type Matrix2D = NonNullable<PptxElement['renderMatrix']>

function multiplyMatrix(left: Matrix2D, right: Matrix2D): Matrix2D {
    const [a, b, c, d, e, f] = left
    const [g, h, i, j, k, l] = right
    return [a * g + c * h, b * g + d * h, a * i + c * j, b * i + d * j, a * k + c * l + e, b * k + d * l + f]
}

function translationMatrix(x: number, y: number): Matrix2D {
    return [1, 0, 0, 1, x, y]
}

function scaleMatrix(x: number, y: number): Matrix2D {
    return [x, 0, 0, y, 0, 0]
}

function rotationMatrix(degrees: number): Matrix2D {
    const radians = (degrees * Math.PI) / 180
    const cosine = Math.cos(radians)
    const sine = Math.sin(radians)
    return [cosine, sine, -sine, cosine, 0, 0]
}

function groupTransformMatrix(group: Element): Matrix2D | undefined {
    const transform = child(child(group, 'grpSpPr'), 'xfrm')
    const offset = child(transform, 'off')
    const extent = child(transform, 'ext')
    const childOffset = child(transform, 'chOff')
    const childExtent = child(transform, 'chExt')
    const width = numberAttr(extent, 'cx')
    const height = numberAttr(extent, 'cy')
    const childWidth = numberAttr(childExtent, 'cx')
    const childHeight = numberAttr(childExtent, 'cy')
    if (!transform || !childOffset || !childExtent || width <= 0 || height <= 0 || childWidth <= 0 || childHeight <= 0) return undefined

    const centerX = width / 2
    const centerY = height / 2
    const flip: Matrix2D = [transform.getAttribute('flipH') === '1' || transform.getAttribute('flipH') === 'true' ? -1 : 1, 0, 0, transform.getAttribute('flipV') === '1' || transform.getAttribute('flipV') === 'true' ? -1 : 1, 0, 0]
    return multiplyMatrix(
        translationMatrix(numberAttr(offset, 'x'), numberAttr(offset, 'y')),
        multiplyMatrix(
            translationMatrix(centerX, centerY),
            multiplyMatrix(
                rotationMatrix(numberAttr(transform, 'rot') / 60_000),
                multiplyMatrix(flip, multiplyMatrix(translationMatrix(-centerX, -centerY), multiplyMatrix(scaleMatrix(width / childWidth, height / childHeight), translationMatrix(-numberAttr(childOffset, 'x'), -numberAttr(childOffset, 'y')))))
            )
        )
    )
}

function applyGroupTransform(element: PptxElement, parent: Matrix2D): PptxElement {
    const transform = multiplyMatrix(
        parent,
        multiplyMatrix(
            translationMatrix(element.x, element.y),
            multiplyMatrix(translationMatrix(element.width / 2, element.height / 2), multiplyMatrix(rotationMatrix(element.rotation), multiplyMatrix([element.flipH ? -1 : 1, 0, 0, element.flipV ? -1 : 1, 0, 0], translationMatrix(-element.width / 2, -element.height / 2))))
        )
    )
    const corners = [
        [0, 0],
        [element.width, 0],
        [element.width, element.height],
        [0, element.height]
    ].map(([x, y]) => [transform[0] * x! + transform[2] * y! + transform[4], transform[1] * x! + transform[3] * y! + transform[5]])
    const minX = Math.min(...corners.map(([x]) => x!))
    const minY = Math.min(...corners.map(([, y]) => y!))
    return {
        ...element,
        x: minX,
        y: minY,
        rotation: 0,
        flipH: false,
        flipV: false,
        motionParent: {matrix: parent, x: element.x, y: element.y},
        renderMatrix: [transform[0], transform[1], transform[2], transform[3], transform[4] - minX, transform[5] - minY]
    }
}

async function parseSceneElements(node: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>, parentTransform?: Matrix2D): Promise<PptxElement[]> {
    if (node.localName === 'grpSp') {
        const localTransform = groupTransformMatrix(node)
        if (!localTransform) {
            warnings.add('A grouped shape with invalid coordinate extents was skipped.')
            return []
        }
        const transform = parentTransform ? multiplyMatrix(parentTransform, localTransform) : localTransform
        const elements: PptxElement[] = []
        for (const item of children(node)) {
            if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(item.localName)) continue
            elements.push(...(await parseSceneElements(item, rels, zip, theme, warnings, urls, imageUrlCache, transform)))
        }
        return elements
    }

    let element: PptxElement | undefined
    if (node.localName === 'sp' || node.localName === 'cxnSp') element = parseShape(node, theme, warnings)
    else if (node.localName === 'pic') element = await parsePicture(node, rels, zip, theme, warnings, urls, imageUrlCache)
    else if (node.localName === 'graphicFrame') {
        element = parseTable(node, theme, warnings)
        if (!element && firstDescendant(node, 'chart')) element = await parseChart(node, rels, zip, theme, warnings)
        else if (!element) {
            element = await parseOlePreview(node, rels, zip, theme, warnings, urls, imageUrlCache)
            if (!element) element = await parseSmartArtTextFallback(node, rels, zip, theme, warnings)
            if (!element) warnings.add('Unsupported graphic frames are not rendered yet.')
        }
    }
    if (element && node.localName === 'cxnSp') warnings.add('Connector geometry is not fully rendered yet.')
    if (!element) return []
    return [parentTransform ? applyGroupTransform(element, parentTransform) : element]
}

async function backgroundFill(xml: XMLDocument | undefined, rels: Map<string, PptxRelationship>, theme: Record<string, string>, warnings: Set<string>, zip: JSZip, urls: string[], imageUrlCache: Map<string, string>): Promise<{color?: string; image?: PptxSlide['backgroundImage']}> {
    const background = xml && firstDescendant(xml.documentElement, 'bg')
    const properties = child(background, 'bgPr')
    if (!properties) {
        if (child(background, 'bgRef')) warnings.add('Theme slide background references are not rendered.')
        return {}
    }
    const imageFill = child(properties, 'blipFill')
    if (!imageFill) {
        if (child(properties, 'grpFill')) warnings.add('Group slide background fills are not rendered.')
        const color = fillColor(properties, theme, warnings)
        return {color: color === 'transparent' ? undefined : color}
    }
    const tile = child(imageFill, 'tile')
    const blip = child(imageFill, 'blip')
    const relationId = namespacedAttr(blip, 'embed') || namespacedAttr(blip, 'link') || ''
    const relation = rels.get(relationId)
    if (!relation || relation.external || namespacedAttr(blip, 'link')) {
        warnings.add('An external or unresolved slide background image was ignored.')
        return {}
    }
    const url = await imageUrlFor(relation.target, zip, urls, imageUrlCache)
    if (!url) {
        warnings.add('A slide background image could not be resolved.')
        return {}
    }
    const sourceRect = child(imageFill, 'srcRect')
    const cropValue = (name: string) => (sourceRect?.hasAttribute(name) ? Number(sourceRect.getAttribute(name)) : 0)
    const crop = {
        left: cropValue('l'),
        top: cropValue('t'),
        right: cropValue('r'),
        bottom: cropValue('b')
    }
    if (Object.values(crop).some(value => !Number.isSafeInteger(value) || value < 0 || value > 100_000) || crop.left + crop.right >= 100_000 || crop.top + crop.bottom >= 100_000) {
        warnings.add('An invalid slide background image crop was ignored.')
        return {}
    }
    if (tile) {
        // DrawingML tile defaults: top-left alignment, no flip, 100% scale, and zero offset.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c0c046ec-a61d-405d-88fe-74d8487a37d7
        const readTileScale = (name: 'sx' | 'sy') => {
            const raw = tile.getAttribute(name)
            if (raw === null) return 1
            const value = raw.trim()
            // ST_Percentage is written as thousandths of a percent or a percent string.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/ff18a37e-9bd7-4338-9c37-1e285b5a5dd2
            const scale = /^(?:\d+\.?\d*|\.\d+)%$/.test(value) ? Number(value.slice(0, -1)) / 100 : /^\d+$/.test(value) ? Number(value) / 100_000 : Number.NaN
            if (!Number.isFinite(scale) || scale <= 0) {
                warnings.add(`An invalid slide background tile ${name} scale was replaced with 100%.`)
                return 1
            }
            return scale
        }
        const scaleX = readTileScale('sx')
        const scaleY = readTileScale('sy')
        const rawAlign = tile.getAttribute('algn') || 'tl'
        const alignValues = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'] as const
        const align = alignValues.includes(rawAlign as (typeof alignValues)[number]) ? (rawAlign as (typeof alignValues)[number]) : 'tl'
        if (align !== rawAlign) warnings.add(`An invalid slide background tile alignment "${rawAlign}" was replaced with top-left.`)
        const readTileOffset = (name: 'tx' | 'ty') => {
            if (!tile.hasAttribute(name)) return 0
            const raw = tile.getAttribute(name) || ''
            const value = Number(raw)
            if (!/^[+-]?\d+$/.test(raw.trim()) || !Number.isSafeInteger(value)) {
                warnings.add(`An invalid slide background tile ${name} offset was replaced with zero.`)
                return 0
            }
            return value
        }
        const rawFlip = tile.getAttribute('flip') || 'none'
        const flips = ['none', 'x', 'y', 'xy'] as const
        const flip = flips.includes(rawFlip as (typeof flips)[number]) ? (rawFlip as (typeof flips)[number]) : 'none'
        const hasCrop = Object.values(crop).some(value => value !== 0)
        if (hasCrop) warnings.add('A cropped tiled slide background uses the full image while tiling.')
        if (flip !== rawFlip) warnings.add(`Unknown slide background tile flip "${rawFlip}" was replaced with none.`)
        return {
            image: {
                url,
                crop: {left: 0, top: 0, right: 0, bottom: 0},
                tile: {scaleX, scaleY, align, flip, offsetX: readTileOffset('tx'), offsetY: readTileOffset('ty')}
            }
        }
    }
    return {image: {url, crop}}
}

async function parseDecorations(path: string, xml: XMLDocument, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>): Promise<PptxElement[]> {
    const tree = firstDescendant(xml.documentElement, 'spTree')
    const rels = await loadRelationships(zip, relationshipPath(path))
    const elements: PptxElement[] = []
    for (const node of children(tree)) {
        if (node.localName === 'sp' && firstDescendant(node, 'ph')) continue
        elements.push(...(await parseSceneElements(node, rels, zip, theme, warnings, urls, imageUrlCache)))
    }
    return elements.map(element => ({...element, id: `${path}:${element.id}`}))
}

export async function parsePptx(file: File): Promise<PptxDocument> {
    if (file.size > MAX_FILE_BYTES) throw new Error('This presentation is larger than the 150 MB browser limit.')
    const signature = new Uint8Array(await file.slice(0, 8).arrayBuffer())
    if (signature[0] === 0xd0 && signature[1] === 0xcf && signature[2] === 0x11 && signature[3] === 0xe0) {
        throw new Error('This file is encrypted or uses the legacy .ppt format. Save it as an unencrypted .pptx first.')
    }
    const urls: string[] = []
    try {
        const zip = await JSZip.loadAsync(file)
        validateArchiveSize(zip)
        const presentationPath = 'ppt/presentation.xml'
        const presentationEntry = zip.file(presentationPath)
        if (!presentationEntry) throw new Error('This file is not a valid PowerPoint presentation.')
        const presentation = parseXml(await presentationEntry.async('string'), presentationPath)
        const presentationRels = await loadRelationships(zip, 'ppt/_rels/presentation.xml.rels')
        const size = firstDescendant(presentation.documentElement, 'sldSz')
        const width = numberAttr(size, 'cx', 12_192_000)
        const height = numberAttr(size, 'cy', 6_858_000)
        const slideIds = firstDescendant(presentation.documentElement, 'sldIdLst')
        const orderedSlides = children(slideIds, 'sldId')
        if (!orderedSlides.length) throw new Error('No slides were found in this presentation.')
        if (orderedSlides.length > 500) throw new Error('This presentation contains more than 500 slides.')
        const fallbackThemePath = Object.keys(zip.files).find(path => /^ppt\/theme\/theme\d+\.xml$/i.test(path))
        const warnings = new Set<string>()
        const slides: PptxSlide[] = []
        const imageUrlCache = new Map<string, string>()
        const themeCache = new Map<string, Record<string, string>>()
        const decorationCache = new Map<string, Promise<PptxElement[]>>()

        const readXml = async (path: string | undefined): Promise<XMLDocument | undefined> => {
            if (!path) return undefined
            const entry = zip.file(path)
            return entry ? parseXml(await entry.async('string'), path) : undefined
        }
        const loadTheme = async (masterRels: Map<string, PptxRelationship>): Promise<Record<string, string>> => {
            const themePath = [...masterRels.values()].find(relation => relation.type.endsWith('/theme'))?.target || fallbackThemePath || 'default'
            const cached = themeCache.get(themePath)
            if (cached) return cached
            const themeXml = await readXml(themePath === 'default' ? undefined : themePath)
            const theme = parseTheme(themeXml)
            themeCache.set(themePath, theme)
            return theme
        }
        const loadDecorations = (path: string | undefined, xml: XMLDocument | undefined, theme: Record<string, string>): Promise<PptxElement[]> => {
            if (!path || !xml) return Promise.resolve([])
            const cacheKey = `${path}|${Object.entries(theme)
                .map(([key, value]) => `${key}:${value}`)
                .join(';')}`
            let cached = decorationCache.get(cacheKey)
            if (!cached) {
                cached = parseDecorations(path, xml, zip, theme, warnings, urls, imageUrlCache)
                decorationCache.set(cacheKey, cached)
            }
            return cached
        }

        for (const [index, slideId] of orderedSlides.entries()) {
            const relId = namespacedAttr(slideId, 'id') || ''
            const slidePath = presentationRels.get(relId)?.target
            if (!slidePath) {
                warnings.add(`Slide ${index + 1} could not be resolved and was skipped.`)
                continue
            }
            const slideEntry = zip.file(slidePath)
            if (!slideEntry) {
                warnings.add(`Slide ${index + 1} is missing from the archive and was skipped.`)
                continue
            }
            const xml = parseXml(await slideEntry.async('string'), slidePath)
            const rels = await loadRelationships(zip, relationshipPath(slidePath))
            const notesRelation = [...rels.values()].find(relation => relation.type.endsWith('/notesSlide'))
            const notesXml = notesRelation && !notesRelation.external ? await readXml(notesRelation.target) : undefined
            if (notesRelation && notesRelation.external) warnings.add(`Slide ${index + 1} has an external notes relationship that was ignored.`)
            else if (notesRelation && !notesXml) warnings.add(`Slide ${index + 1} references a missing notes slide.`)
            const layoutPath = [...rels.values()].find(relation => relation.type.endsWith('/slideLayout'))?.target
            const layoutXml = await readXml(layoutPath)
            if (layoutPath && !layoutXml) warnings.add(`Slide ${index + 1} references a missing layout.`)
            const layoutRels = layoutPath ? await loadRelationships(zip, relationshipPath(layoutPath)) : new Map<string, PptxRelationship>()
            const masterPath = [...layoutRels.values()].find(relation => relation.type.endsWith('/slideMaster'))?.target
            const masterXml = await readXml(masterPath)
            if (masterPath && !masterXml) warnings.add(`Slide ${index + 1} references a missing master.`)
            const masterRels = masterPath ? await loadRelationships(zip, relationshipPath(masterPath)) : new Map<string, PptxRelationship>()
            const theme = await loadTheme(masterRels)
            const masterTextStyles = parseMasterTextStyles(masterXml, theme, warnings)
            const masterPlaceholders = parsePlaceholderCandidates(masterXml, theme, warnings, masterTextStyles)
            const layoutPlaceholders = parsePlaceholderCandidates(layoutXml, theme, warnings, masterTextStyles, masterPlaceholders)
            const tree = firstDescendant(xml.documentElement, 'spTree')
            const masterElements = await loadDecorations(masterPath, masterXml, theme)
            const layoutElements = await loadDecorations(layoutPath, layoutXml, theme)
            const elements: PptxElement[] = [...masterElements, ...layoutElements]
            for (const node of children(tree)) {
                if (node.localName === 'sp') {
                    const identity = parsePlaceholderIdentity(node)
                    const layoutPlaceholder = matchPlaceholder(identity, layoutPlaceholders)
                    const masterPlaceholder = matchPlaceholder(identity, masterPlaceholders)
                    if (!identity) {
                        elements.push(parseShape(node, theme, warnings))
                        continue
                    }
                    const masterStyle = masterTextStyles[placeholderStyleCategory(identity.type)]
                    const inheritedParagraphStyles = layoutPlaceholder?.paragraphStyles || masterPlaceholder?.paragraphStyles || {}
                    let element = parseShape(node, theme, warnings, masterStyle?.style || {}, inheritedParagraphStyles)
                    element = inheritPlaceholderProperties(element, node, layoutPlaceholder || masterPlaceholder)
                    elements.push(element)
                } else if (['pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(node.localName)) {
                    elements.push(...(await parseSceneElements(node, rels, zip, theme, warnings, urls, imageUrlCache)))
                }
            }
            const animations = parseAnimations(xml.documentElement, theme, warnings, elements, width, height)
            const slideInfo = firstDescendant(xml.documentElement, 'cSld')
            const slideBackground = await backgroundFill(xml, rels, theme, warnings, zip, urls, imageUrlCache)
            const layoutBackground = await backgroundFill(layoutXml, layoutRels, theme, warnings, zip, urls, imageUrlCache)
            const masterBackground = await backgroundFill(masterXml, masterRels, theme, warnings, zip, urls, imageUrlCache)
            const selectedBackground = slideBackground.image || slideBackground.color ? slideBackground : layoutBackground.image || layoutBackground.color ? layoutBackground : masterBackground
            slides.push({
                id: relId || `slide-${index + 1}`,
                name: child(slideInfo, 'name')?.getAttribute('val') || `Slide ${index + 1}`,
                elements,
                speakerNotes: parseSpeakerNotes(notesXml, theme, warnings),
                ...animations,
                ...parseSlideAdvance(xml.documentElement, warnings),
                background: selectedBackground.color || '#ffffff',
                backgroundImage: selectedBackground.image,
                transition: parseTransition(xml.documentElement, warnings)
            })
        }
        if (!slides.length) throw new Error('No readable slides were found in this presentation.')
        return {name: file.name, width, height, slides, warnings: [...warnings], objectUrls: urls}
    } catch (error) {
        urls.forEach(URL.revokeObjectURL)
        const message = error instanceof Error ? error.message : 'The presentation could not be read.'
        if (/encrypt|password/i.test(message)) throw new Error('Password-protected presentations are not supported.')
        if (/zip|central directory|unexpected end/i.test(message)) throw new Error('The PPTX file appears damaged or incomplete.')
        throw error instanceof Error ? error : new Error(message)
    }
}
