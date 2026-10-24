import JSZip from 'jszip'
import presetShapeDefinitionsUrl from './preset-shape-definitions.xml?url'
import { normalizePptxError, validatePptxFile } from './pptx-input'
import {inverseTimingProgress, parseAnimationSpeed, parseOpenXmlPercentage, parseTimingFilter, type TimingCurve} from './animation-timing'
import {animationGeometryVariables, evaluateAnimationFormula as evaluateMotionFormula, sampleAnimationFormula} from './animation-formula'
import {createTextIterationSegmenter, textIterationLocale, textIterationRanges} from './text-iteration'
import {getChartDateAxisMajorTicks, isSupportedChartNumberFormat} from './chart-number-format'
import {hslToRgb} from './color'
import {getPptxChartStacking, isSvgPolygonGeometry, presetEllipsePoint} from './pptx-render-utils'
import type {
    AnimationOnOffProperty,
    AnimationShadowOffsetProperty,
    AnimatedLineArrow,
    AnimatedLineArrowProperty,
    AnimatedLineArrowSize,
    AnimatedLineArrowSizeProperty,
    AnimatedStrokeDashStyle,
    AnimatedVisibility,
    PptxAnimation,
    PptxAnimationTriggerCondition,
    PptxAlphaInverse,
    PptxChart,
    PptxChartAxis,
    PptxChartAxisCrosses,
    PptxChartAxisPosition,
    PptxChartDataLabelPosition,
    PptxChartDataLabelStyle,
    PptxChartGrouping,
    PptxChartMarkerSymbol,
    PptxChartSeries,
    PptxChartTickLabelPosition,
    PptxChartTickMark,
    PptxCustomPath,
    PptxDocument,
    PptxElement,
    PptxEmbeddedFontFace,
    PptxGlow,
    PptxHyperlink,
    PptxLineEnd,
    PptxParagraph,
    PptxRun,
    PptxShadow,
    PptxSlide,
    PptxTableCell,
    PptxTextOrientation,
    TextStyleAnimationProperty,
    TextStyleAnimationValue
} from './pptx-model'

// ponytail: cap animated strokes at about 10,500 CSS px before stage scaling; raise only for real decks that need wider strokes.
const MAX_ANIMATED_STROKE_WIDTH_EMU = 100_000_000
const MAX_ANIMATED_SHADOW_OFFSET_EMU = 100_000_000
const FLY_PRESET_DIRECTIONS = new Map<number, NonNullable<PptxAnimation['slideFrom']>>([
    [1, 'top'], [2, 'right'], [4, 'bottom'], [8, 'left'],
    [9, 'topLeft'], [3, 'topRight'], [12, 'bottomLeft'], [6, 'bottomRight']
])
const PEEK_PRESET_DIRECTIONS = new Map<number, NonNullable<PptxAnimation['slideFrom']>>([
    [1, 'top'], [2, 'right'], [4, 'bottom'], [8, 'left']
])
const STRETCH_PRESET_DIRECTIONS = new Map<number, {axis: 'x' | 'y'; origin: [number, number]}>([
    // Stretch grows away from the edge opposite its named direction.
    [1, {axis: 'y', origin: [0.5, 1]}],
    [2, {axis: 'x', origin: [0, 0.5]}],
    [4, {axis: 'y', origin: [0.5, 0]}],
    [8, {axis: 'x', origin: [1, 0.5]}],
    [10, {axis: 'x', origin: [0.5, 0.5]}]
])
const SWIVEL_PRESET_AXES = new Map<number, NonNullable<PptxAnimation['rotationAxis']>>([
    [10, 'y'], // Horizontal swivel turns around the vertical axis.
    [5, 'x'] // Vertical swivel turns around the horizontal axis.
])
// The Open XML SDK exposes presetSubtype as Int32; these documented Wipe/Wheel codes are compatibility mappings.
// Explicit out-of-table values fall back with a warning instead of being silently clamped to another preset.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode.presetsubtype?view=openxml-3.0.1
const WIPE_PRESET_DIRECTIONS = new Map<number, NonNullable<PptxAnimation['wipeDirection']>>([
    [1, 'up'], [2, 'right'], [4, 'down'], [8, 'left']
])
const WHEEL_PRESET_SPOKES = new Map<number, NonNullable<PptxAnimation['wheelSpokes']>>([
    [1, 1], [2, 2], [3, 3], [4, 4], [8, 8]
])
const ZOOM_PRESET_SUBTYPES = new Map<number, {direction: 'in' | 'out'; from: number; to: number; origin?: 'screenCenter' | 'screenBottom'; approximate?: boolean}>([
    [16, {direction: 'in', from: 0, to: 1}],
    [32, {direction: 'out', from: 1, to: 0}],
    [528, {direction: 'in', from: 0, to: 1, origin: 'screenCenter'}],
    [36, {direction: 'out', from: 1, to: 0, origin: 'screenBottom'}],
    [288, {direction: 'out', from: 1, to: 0.8, approximate: true}],
    [272, {direction: 'in', from: 0.8, to: 1, approximate: true}]
])
// OOXML defines presetSubtype as an integer bitflag whose meaning depends on the effect.
// The Split codes below follow the legacy MS-PPT direction table as an Office-compatibility inference.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode.presetsubtype?view=openxml-3.0.1
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
const SPLIT_PRESET_DIRECTIONS = new Map<number, {orientation: NonNullable<PptxAnimation['barnOrientation']>; motion: NonNullable<PptxAnimation['barnMotion']>}>([
    [21, {orientation: 'vertical', motion: 'in'}],
    [37, {orientation: 'vertical', motion: 'out'}],
    [26, {orientation: 'horizontal', motion: 'in'}],
    [42, {orientation: 'horizontal', motion: 'out'}]
])
const SHAPE_ANIMATION_PRESET_IDS = new Map<number, NonNullable<PptxAnimation['shapeFilter']>>([
    [4, 'box'], [6, 'circle'], [8, 'diamond'], [13, 'plus']
])
// OOXML p:cTn@presetSubtype is an effect-specific integer bitflag, not the MsoAnimDirection enum.
// The legacy MS-PPT TimeEffectID values below are used only as an Office-compatibility inference
// when animEffect.filter is absent or unsupported; a supported named OOXML filter takes precedence.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
const PRESET_BLINDS_ORIENTATIONS = new Map<number, NonNullable<PptxAnimation['blindsOrientation']>>([[5, 'vertical'], [10, 'horizontal']])
const PRESET_CHECKER_ORIENTATIONS = new Map<number, NonNullable<PptxAnimation['checkerOrientation']>>([[5, 'vert'], [10, 'horz']])
const PRESET_RANDOM_BAR_ORIENTATIONS = new Map<number, NonNullable<PptxAnimation['randomBarOrientation']>>([[5, 'vertical'], [10, 'horizontal']])
const PRESET_STRIPS_DIRECTIONS = new Map<number, NonNullable<PptxAnimation['stripsDirection']>>([[9, 'lu'], [3, 'ru'], [6, 'rd'], [12, 'ld']])

function presetVariantOrDefault<T>(subtype: number | undefined, rawSubtype: string, wasSpecified: boolean, variants: Map<number, T>, defaultValue: T, presetName: string, warnings: Set<string>): T {
    if (!wasSpecified || subtype === 0) return defaultValue
    const variant = subtype === undefined ? undefined : variants.get(subtype)
    if (variant !== undefined) return variant
    const subtypeIssue = subtype === undefined ? 'is invalid' : 'is unsupported'
    warnings.add(`The ${presetName} preset subtype "${rawSubtype}" ${subtypeIssue} without a filter string; its default direction was used.`)
    return defaultValue
}

const ANIMATION_ATTRIBUTE_NAMES = new Map([
    'ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ScaleX', 'ScaleY',
    'style.fontSize', 'style.fontWeight', 'style.fontStyle', 'style.fontFamily', 'style.textDecorationUnderline',
    'style.textDecorationLineThrough', 'style.textTransform', 'style.textEffectOutline', 'style.textShadow',
    'style.textEffectEmboss', 'style.visibility', 'style.rotation', 'style.opacity', 'style.color',
    'imageData.chromakey', 'imageData.grayscale', 'imageData.cropLeft', 'imageData.cropTop', 'imageData.cropRight',
    'imageData.cropBottom', 'imageData.gain', 'imageData.blacklevel', 'imageData.blackleve', 'imageData.gamma',
    'fill.on', 'fill.opacity', 'fillcolor', 'stroke.on', 'stroke.dashstyle', 'stroke.startArrow', 'stroke.endArrow',
    'stroke.startArrowWidth', 'stroke.endArrowWidth', 'stroke.startArrowLength', 'stroke.endArrowLength',
    'stroke.opacity', 'stroke.weight', 'stroke.color', 'shadow.on', 'shadow.opacity', 'shadow.color',
    'shadow.offset.x', 'shadow.offset.y'
].map(name => [name.toLowerCase(), name]))
// Generic p:anim uses the same renderer for Office's ppt_r/r angle attributes.
// Formula evaluation still exposes #ppt_r and #r as the original property baselines.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/be583d68-11d7-43c2-a300-8ef2ffac7841
ANIMATION_ATTRIBUTE_NAMES.set('ppt_r', 'style.rotation')
ANIMATION_ATTRIBUTE_NAMES.set('r', 'style.rotation')
ANIMATION_ATTRIBUTE_NAMES.set('fill.color', 'fillcolor')

function normalizeAnimationAttributeName(value: string | null | undefined): string | undefined {
    const name = value?.trim()
    if (!name) return undefined
    return ANIMATION_ATTRIBUTE_NAMES.get(name.toLowerCase()) || name
}

// The raw OOXML subtype mapping follows the supplied parser reference; Microsoft's Open XML API exposes this attribute as an integer.
function getShapePresetDirection(transition: 'in' | 'out', subtype?: number): 'in' | 'out' {
    if (subtype === 16) return transition === 'in' ? 'in' : 'out'
    if (subtype === 32) return transition === 'in' ? 'out' : 'in'
    return transition
}
const PRESET_GEOMETRIES_WITH_HANDWRITTEN_RENDERERS = new Set(['rect', 'roundRect', 'round1Rect', 'round2DiagRect', 'round2SameRect', 'snipRoundRect', 'ellipse', 'line', 'lineInv', 'straightConnector1', 'plus', 'mathMinus', 'mathEqual', 'mathPlus', 'mathDivide', 'mathMultiply', 'mathNotEqual', 'snip1Rect', 'snip2DiagRect', 'snip2SameRect', 'homePlate', 'nonIsoscelesTrapezoid', 'heptagon', 'decagon', 'leftRightArrow', 'upDownArrow', 'upArrow', 'downArrow', 'bentUpArrow', 'quadArrow', 'leftRightUpArrow', 'rightArrow', 'leftArrow', 'notchedRightArrow', 'heart', 'lightningBolt', 'triangle', 'rtTriangle', 'diamond', 'cube', 'plaque', 'frame', 'halfFrame', 'cornerTabs', 'squareTabs', 'plaqueTabs', 'donut', 'pie', 'chord', 'arc', 'blockArc', 'pieWedge', 'flowChartConnector', 'flowChartSummingJunction', 'flowChartOr', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDelay', 'flowChartDisplay', 'flowChartInternalStorage', 'flowChartDecision', 'flowChartSort', 'flowChartCollate', 'flowChartPunchedCard', 'flowChartOfflineStorage', 'flowChartOnlineStorage', 'flowChartPunchedTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartMagneticTape', 'flowChartMultidocument', 'flowChartOffpageConnector', 'flowChartDocument', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartInputOutput', 'flowChartPreparation', 'flowChartTerminator', 'flowChartPredefinedProcess', 'flowChartExtract', 'flowChartMerge', 'parallelogram', 'trapezoid', 'hexagon', 'octagon', 'pentagon', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32', 'dodecagon', 'chevron', 'leftBracket', 'rightBracket', 'bracketPair', 'leftBrace', 'rightBrace', 'bracePair', 'custom', 'cloud'])

interface PptxRelationship {
    target: string
    type: string
    external?: boolean
}

type PptxTextStyle = Omit<PptxRun, 'text'>

type ParagraphValues = Partial<Omit<PptxParagraph, 'runs'>>

interface ParagraphDefaults {
    values: ParagraphValues
    text: PptxTextStyle
    tabStops?: {position: number; alignment?: 'left' | 'center' | 'right' | 'decimal'}[]
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

const MAX_ARCHIVE_BYTES = 250 * 1024 * 1024
const MAX_ARCHIVE_ENTRIES = 20_000
const MAX_CHART_WORKBOOK_BYTES = 32 * 1024 * 1024
const MAX_CHART_WORKSHEET_BYTES = 8 * 1024 * 1024
const MAX_CHART_SHARED_STRINGS_BYTES = 8 * 1024 * 1024
const MAX_CHART_REFERENCES = 512
const MAX_CHART_REFERENCED_CELLS = 32_768
const LINE_GEOMETRIES = new Set(['line', 'lineInv', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5', 'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5'])
// MS-CFB compound-file header signature: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea
const MAX_EMBEDDED_FONT_BYTES = 64 * 1024 * 1024
const MAX_EMBEDDED_FONT_FACES = 256
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

// ECMA-376 Part 1, presetShapeDefinitions.xml, cloud.
let cloudPresetGeometryCache: Element | undefined
function getCloudPresetGeometry(): Element {
    if (!cloudPresetGeometryCache) {
        cloudPresetGeometryCache = parseXml(`<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:gdLst>
                <a:gd name="il" fmla="*/ w 2977 21600"/><a:gd name="it" fmla="*/ h 3262 21600"/>
                <a:gd name="ir" fmla="*/ w 17087 21600"/><a:gd name="ib" fmla="*/ h 17337 21600"/>
                <a:gd name="g27" fmla="*/ w 67 21600"/><a:gd name="g28" fmla="*/ h 21577 21600"/>
                <a:gd name="g29" fmla="*/ w 21582 21600"/><a:gd name="g30" fmla="*/ h 1235 21600"/>
            </a:gdLst>
            <a:rect l="il" t="it" r="ir" b="ib"/>
            <a:pathLst>
                <a:path w="43200" h="43200">
                    <a:moveTo><a:pt x="3900" y="14370"/></a:moveTo>
                    <a:arcTo wR="6753" hR="9190" stAng="-11429249" swAng="7426832"/>
                    <a:arcTo wR="5333" hR="7267" stAng="-8646143" swAng="5396714"/>
                    <a:arcTo wR="4365" hR="5945" stAng="-8748475" swAng="5983381"/>
                    <a:arcTo wR="4857" hR="6595" stAng="-7859164" swAng="7034504"/>
                    <a:arcTo wR="5333" hR="7273" stAng="-4722533" swAng="6541615"/>
                    <a:arcTo wR="6775" hR="9220" stAng="-2776035" swAng="7816140"/>
                    <a:arcTo wR="5785" hR="7867" stAng="37501" swAng="6842000"/>
                    <a:arcTo wR="6752" hR="9215" stAng="1347096" swAng="6910353"/>
                    <a:arcTo wR="7720" hR="10543" stAng="3974558" swAng="4542661"/>
                    <a:arcTo wR="4360" hR="5918" stAng="-16496525" swAng="8804134"/>
                    <a:arcTo wR="4345" hR="5945" stAng="-14809710" swAng="9151131"/>
                    <a:close/>
                </a:path>
                <a:path w="43200" h="43200" fill="none" extrusionOk="false">
                    <a:moveTo><a:pt x="4693" y="26177"/></a:moveTo><a:arcTo wR="4345" hR="5945" stAng="5204520" swAng="1585770"/>
                    <a:moveTo><a:pt x="6928" y="34899"/></a:moveTo><a:arcTo wR="4360" hR="5918" stAng="4416628" swAng="686848"/>
                    <a:moveTo><a:pt x="16478" y="39090"/></a:moveTo><a:arcTo wR="6752" hR="9215" stAng="8257449" swAng="844866"/>
                    <a:moveTo><a:pt x="28827" y="34751"/></a:moveTo><a:arcTo wR="6752" hR="9215" stAng="387196" swAng="959901"/>
                    <a:moveTo><a:pt x="34129" y="22954"/></a:moveTo><a:arcTo wR="5785" hR="7867" stAng="-4217541" swAng="4255042"/>
                    <a:moveTo><a:pt x="41798" y="15354"/></a:moveTo><a:arcTo wR="5333" hR="7273" stAng="1819082" swAng="1665090"/>
                    <a:moveTo><a:pt x="38324" y="5426"/></a:moveTo><a:arcTo wR="4857" hR="6595" stAng="-824660" swAng="891534"/>
                    <a:moveTo><a:pt x="29078" y="3952"/></a:moveTo><a:arcTo wR="4857" hR="6595" stAng="-8950887" swAng="1091722"/>
                    <a:moveTo><a:pt x="22141" y="4720"/></a:moveTo><a:arcTo wR="4365" hR="5945" stAng="-9809656" swAng="1061181"/>
                    <a:moveTo><a:pt x="14000" y="5192"/></a:moveTo><a:arcTo wR="6753" hR="9190" stAng="-4002417" swAng="739161"/>
                    <a:moveTo><a:pt x="4127" y="15789"/></a:moveTo><a:arcTo wR="6753" hR="9190" stAng="9459261" swAng="711490"/>
                </a:path>
            </a:pathLst>
        </a:custGeom>`, 'ECMA-376 cloud preset geometry').documentElement
    }
    return cloudPresetGeometryCache
}

// ECMA-376 Part 1's preset definitions supply paths and text rectangles for presets without custom renderers.
let presetShapeDefinitionCache: Map<string, Element> | undefined
let presetShapeDefinitionLoad: Promise<void> | undefined
async function loadPresetShapeDefinitions(): Promise<boolean> {
    if (presetShapeDefinitionCache) return true
    const load = presetShapeDefinitionLoad ??= (async () => {
        const response = await fetch(presetShapeDefinitionsUrl)
        if (!response.ok) throw new Error(`Preset geometry resource returned ${response.status}.`)
        const root = parseXml((await response.text()).replace(/^\uFEFF/, ''), 'ECMA-376 preset shape definitions').documentElement
        presetShapeDefinitionCache = new Map(children(root).map(definition => [definition.localName, definition]))
    })()
    try {
        await load
        return true
    } catch {
        if (presetShapeDefinitionLoad === load) presetShapeDefinitionLoad = undefined
        return false
    }
}

function getPresetShapeDefinition(name: string, source: Element | undefined, width: number, height: number): Element | undefined {
    const definition = presetShapeDefinitionCache?.get(name)
    if (!definition) return undefined

    const result = definition.cloneNode(true) as Element
    for (const path of children(child(result, 'pathLst'), 'path')) {
        if (!path.hasAttribute('w')) path.setAttribute('w', String(width))
        if (!path.hasAttribute('h')) path.setAttribute('h', String(height))
    }
    const sourceAdjustments = children(child(source, 'avLst'), 'gd')
    if (!sourceAdjustments.length) return result

    let adjustmentList = child(result, 'avLst')
    if (!adjustmentList) {
        adjustmentList = result.ownerDocument.createElementNS(DRAWING_NS, 'a:avLst')
        result.insertBefore(adjustmentList, child(result, 'gdLst') ?? null)
    }
    for (const sourceAdjustment of sourceAdjustments) {
        const adjustmentName = sourceAdjustment.getAttribute('name')
        const formula = sourceAdjustment.getAttribute('fmla')
        if (!adjustmentName || !formula) continue
        const defaultAdjustment = children(adjustmentList, 'gd').find(node => node.getAttribute('name') === adjustmentName)
        if (defaultAdjustment) defaultAdjustment.setAttribute('fmla', formula)
        else adjustmentList.appendChild(result.ownerDocument.importNode(sourceAdjustment, true))
    }
    return result
}

function numberAttr(node: Element | undefined, key: string, fallback = 0): number {
    const raw = node?.getAttribute(key)
    if (!raw) return fallback
    const value = Number(raw)
    return Number.isFinite(value) ? value : fallback
}

function parseOfficeAnimationNumber(raw: string | null | undefined): number | undefined {
    const value = raw?.trim()
    if (!value || value.length > 16_384) return undefined
    const decimal = Number(value)
    if (Number.isFinite(decimal)) return decimal

    // MS-OI29500 p:anim values may use radix-prefixed hexadecimal, octal, or binary real numbers.
    // JavaScript Number() accepts their unsigned integers but not signed or fractional forms such as -0X5FE5 and 0x4A.F.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/0b35ab0b-2029-4180-8a76-8d784dac1d2b
    const match = /^([+-]?)(0[xob])([0-9a-f]*)(?:\.([0-9a-f]*))?(?:[eE]([+-]?\d+))?$/i.exec(value)
    if (!match) return undefined
    const radix = match[2]!.toLowerCase() === '0x' ? 16 : match[2]!.toLowerCase() === '0o' ? 8 : 2
    const integerDigits = match[3] || ''
    const fractionalDigits = match[4] || ''
    if (!integerDigits && !fractionalDigits) return undefined
    const digitValue = (character: string) => Number.parseInt(character, 16)
    if ([...integerDigits, ...fractionalDigits].some(character => digitValue(character) >= radix)) return undefined

    let magnitude = 0
    for (const character of integerDigits) magnitude = magnitude * radix + digitValue(character)
    for (let index = 0; index < fractionalDigits.length; index++) {
        magnitude += digitValue(fractionalDigits[index]!) / radix ** (index + 1)
    }
    if (match[5]) magnitude *= 10 ** Number(match[5])
    const result = (match[1] === '-' ? -1 : 1) * magnitude
    return Number.isFinite(result) ? result : undefined
}

function hasValidAnimateValueCombination(hasFrom: boolean, hasTo: boolean, hasBy: boolean): boolean {
    // MS-OI29500 §19.5.1(h): from/to, from/by, to-only, and by-only are valid; when to is present, by is ignored.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
    // Rotation behaviors use the same combinations under MS-OI29500 §19.5.5(b).
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/39903f1a-fcb4-4aee-81e2-66f4fc7493fd
    return (hasFrom && hasTo)
        || (hasFrom && !hasTo && hasBy)
        || (!hasFrom && hasTo)
        || (!hasFrom && !hasTo && hasBy)
}

function percentageValue(value: string | null): number | undefined {
    return parseOpenXmlPercentage(value)
}

function trimXmlWhitespace(value: string | null | undefined): string | undefined {
    return value?.replace(/^[\t\n\r ]+|[\t\n\r ]+$/g, '')
}

function fixedPercentage(value: string | null): number | undefined {
    const percent = percentageValue(value)
    return percent !== undefined && percent >= 0 && percent <= 1 ? percent : undefined
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

interface ChartWorkbookMetadata {
    indexedColors?: string[]
    numberFormatsByFormula: Map<string, Array<string | undefined>>
    valuesByFormula: Map<string, Array<string | null>>
}

const CHART_WORKBOOK_BUILTIN_FORMATS = new Map<number, string>([
    [0, 'General'], [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'],
    [9, '0%'], [10, '0.00%'], [11, '0.00E+00'], [12, '# ?/?'], [13, '# ??/??'],
    [14, 'mm-dd-yy'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'], [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'],
    [20, 'h:mm'], [21, 'h:mm:ss'], [22, 'm/d/yy h:mm'], [37, '#,##0;(#,##0)'], [38, '#,##0;[Red](#,##0)'],
    [39, '#,##0.00;(#,##0.00)'], [40, '#,##0.00;[Red](#,##0.00)'], [45, 'mm:ss'], [46, '[h]:mm:ss'],
    [47, 'mmss.0'], [48, '##0.0E+0'], [49, '@']
])

// shortcut: resolves only one-dimensional A1 ranges; upgrade when unions or external workbook references need support.
function chartWorksheetRange(formula: string): {sheetName: string; cells: string[]} | undefined {
    const reference = formula.trim().replace(/^=/, '')
    const separator = reference.lastIndexOf('!')
    if (separator <= 0) return undefined
    let sheetName = reference.slice(0, separator)
    if (sheetName.includes('[') || sheetName.includes(']')) return undefined
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1).replaceAll("''", "'")
    const range = reference.slice(separator + 1).replaceAll('$', '')
    const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(range)
    if (!sheetName || !match) return undefined
    const columnIndex = (column: string) => [...column.toUpperCase()].reduce((index, letter) => index * 26 + letter.charCodeAt(0) - 64, 0)
    const columnName = (index: number) => {
        let name = ''
        for (let value = index; value > 0; value = Math.floor((value - 1) / 26)) name = String.fromCharCode(65 + ((value - 1) % 26)) + name
        return name
    }
    const startColumn = columnIndex(match[1]!)
    const endColumn = columnIndex(match[3] || match[1]!)
    const startRow = Number(match[2])
    const endRow = Number(match[4] || match[2])
    if (startColumn > 16_384 || endColumn > 16_384 || startRow < 1 || endRow < 1 || startRow > 1_048_576 || endRow > 1_048_576 || (startColumn !== endColumn && startRow !== endRow)) return undefined
    const steps = Math.max(Math.abs(endColumn - startColumn), Math.abs(endRow - startRow)) + 1
    if (steps > 128) return undefined
    const columnStep = Math.sign(endColumn - startColumn)
    const rowStep = Math.sign(endRow - startRow)
    const cells = Array.from({length: steps}, (_, index) => `${columnName(startColumn + index * columnStep)}${startRow + index * rowStep}`)
    return {sheetName, cells}
}

async function chartWorkbookMetadata(chartRoot: Element, chartPath: string, zip: JSZip, warnings: Set<string>): Promise<ChartWorkbookMetadata | undefined> {
    const relationshipId = namespacedAttr(child(chartRoot, 'externalData'), 'id')
    if (!relationshipId) return undefined

    try {
        const workbookRelation = (await loadRelationships(zip, relationshipPath(chartPath))).get(relationshipId)
        if (!workbookRelation) {
            warnings.add('An embedded chart workbook relationship is missing; indexed colors use the default palette and source-linked formats use chart caches.')
            return undefined
        }
        if (workbookRelation.external) {
            warnings.add('An external chart workbook is not read; indexed colors use the default palette and source-linked cell styles are unavailable.')
            return undefined
        }
        const workbookPart = zip.file(workbookRelation.target)
        if (!workbookPart) {
            warnings.add('An embedded chart workbook is missing; indexed colors use the default palette and source-linked formats use chart caches.')
            return undefined
        }

        const workbookSize = (workbookPart as JSZip.JSZipObject & {_data?: {uncompressedSize?: number}})._data?.uncompressedSize
        if (typeof workbookSize !== 'number' || !Number.isFinite(workbookSize) || workbookSize < 0 || workbookSize > MAX_CHART_WORKBOOK_BYTES) {
            warnings.add('An embedded chart workbook size is unavailable or exceeds the 32 MB browser reading limit; chart-cache values are used and source-linked formats fall back to raw values.')
            return undefined
        }
        const workbook = await JSZip.loadAsync(await workbookPart.async('uint8array'))
        await validateArchiveSize(workbook)
        const packageRelationships = workbook.file('_rels/.rels')
        if (!packageRelationships) {
            warnings.add('An embedded chart workbook has no package relationships; source-linked cell styles are unavailable.')
            return undefined
        }
        const packageXml = parseXml(await packageRelationships.async('string'), workbookRelation.target)
        const workbookTarget = descendants(packageXml.documentElement, 'Relationship')
            .find(item => item.getAttribute('Type')?.endsWith('/officeDocument') && item.getAttribute('TargetMode') !== 'External')
            ?.getAttribute('Target')
        if (!workbookTarget) {
            warnings.add('An embedded chart workbook has no workbook part; source-linked cell styles are unavailable.')
            return undefined
        }

        const workbookPath = resolveTarget('', workbookTarget)
        const workbookXmlPart = workbook.file(workbookPath)
        if (!workbookXmlPart) return undefined
        const workbookXml = parseXml(await workbookXmlPart.async('string'), workbookPath)
        const worksheetNames = children(child(workbookXml.documentElement, 'sheets'), 'sheet').map(sheet => sheet.getAttribute('name') || '')
        const definedNames = children(child(workbookXml.documentElement, 'definedNames'), 'definedName')
        const workbookRelationships = await loadRelationships(workbook, relationshipPath(workbookPath))
        const stylesRelation = [...workbookRelationships.values()]
            .find(item => item.type.endsWith('/styles') && !item.external)
        const stylesPart = stylesRelation ? workbook.file(stylesRelation.target) : undefined
        if (!stylesPart) warnings.add('An embedded chart workbook styles part is missing; source-linked cell styles are unavailable.')
        const stylesXml = stylesPart ? parseXml(await stylesPart.async('string'), stylesRelation!.target) : undefined
        const indexedColors = stylesXml && firstDescendant(stylesXml.documentElement, 'indexedColors')
        const customFormats = new Map<number, string>()
        for (const item of children(firstDescendant(stylesXml?.documentElement, 'numFmts'), 'numFmt')) {
            const id = Number(item.getAttribute('numFmtId'))
            const code = item.getAttribute('formatCode')
            if (Number.isSafeInteger(id) && id >= 0 && code) customFormats.set(id, code)
        }
        const styleFormats = children(firstDescendant(stylesXml?.documentElement, 'cellXfs'), 'xf')
        const baseStyleFormats = children(firstDescendant(stylesXml?.documentElement, 'cellStyleXfs'), 'xf')
        const styleNumberFormat = (styleIndex: string | null): string | undefined => {
            const index = styleIndex == null ? 0 : Number(styleIndex)
            const cellFormat = Number.isSafeInteger(index) && index >= 0 ? styleFormats[index] : undefined
            if (!cellFormat) return undefined
            const directId = Number(cellFormat.getAttribute('numFmtId') || '0')
            const baseIndex = Number(cellFormat.getAttribute('xfId') || '0')
            const baseId = Number.isSafeInteger(baseIndex) && baseIndex >= 0 ? Number(baseStyleFormats[baseIndex]?.getAttribute('numFmtId') || '0') : 0
            const apply = cellFormat.getAttribute('applyNumberFormat')?.toLowerCase()
            const numberFormatId = ['0', 'false'].includes(apply || '') ? baseId : directId !== 0 || ['1', 'true'].includes(apply || '') ? directId : baseId
            const formatCode = customFormats.get(numberFormatId) || CHART_WORKBOOK_BUILTIN_FORMATS.get(numberFormatId)
            if (!formatCode || formatCode.toLowerCase() === 'general') return undefined
            if (isSupportedChartNumberFormat(formatCode)) return formatCode
            warnings.add('An embedded chart workbook number format is unsupported and its chart values use raw numbers.')
            return undefined
        }
        const sharedStringsRelation = [...workbookRelationships.values()].find(item => item.type.endsWith('/sharedStrings') && !item.external)
        const sharedStringsPart = sharedStringsRelation ? workbook.file(sharedStringsRelation.target) : undefined
        let sharedStrings: string[] = []
        if (sharedStringsPart && descendants(chartRoot, 'strRef').some(reference => !!child(reference, 'f') && !chartCache(reference, warnings).some(value => value !== null))) {
            const sharedStringsSize = (sharedStringsPart as JSZip.JSZipObject & {_data?: {uncompressedSize?: number}})._data?.uncompressedSize
            if (sharedStringsSize === undefined || sharedStringsSize > MAX_CHART_SHARED_STRINGS_BYTES) {
                warnings.add('Embedded chart workbook shared strings are unavailable or exceed the 8 MB reading limit; uncached category values may be omitted.')
            } else {
                try {
                    const sharedStringsXml = parseXml(await sharedStringsPart.async('string'), sharedStringsRelation!.target)
                    sharedStrings = children(sharedStringsXml.documentElement, 'si').map(item => descendants(item, 't').map(text => text.textContent || '').join(''))
                } catch {
                    warnings.add('Embedded chart workbook shared strings could not be read; uncached category values may be omitted.')
                }
            }
        }
        const rangesByFormula = new Map<string, {sheetName: string; cells: string[]}>()
        // shortcut: resolves alias chains to one-dimensional A1 ranges with cycle guards; upgrade when computed names need support.
        const chartRange = (formula: string, scopeSheetIndex?: number, fromDefinedName = false, visited = new Set<string>(), depth = 0): {sheetName: string; cells: string[]} | undefined => {
            const direct = chartWorksheetRange(formula)
            if (direct) return direct
            if (depth >= 16) return undefined
            const reference = formula.trim().replace(/^=/, '')
            if (reference.includes('[') || reference.includes(']')) return undefined
            const separator = reference.lastIndexOf('!')
            const sheetReference = separator > 0 ? reference.slice(0, separator) : undefined
            let sheetName = sheetReference
            if (sheetName?.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1).replaceAll("''", "'")
            const name = separator > 0 ? reference.slice(separator + 1) : reference
            if (!name) return undefined
            const sheetIndex = sheetName === undefined ? scopeSheetIndex : worksheetNames.findIndex(item => item.toLowerCase() === sheetName.toLowerCase())
            if ((sheetName !== undefined && sheetIndex === -1) || (sheetIndex !== undefined && !worksheetNames[sheetIndex])) return undefined
            if (fromDefinedName && sheetIndex !== undefined) {
                const localSheetName = worksheetNames[sheetIndex]!
                const relativeRange = chartWorksheetRange(`'${localSheetName.replaceAll("'", "''")}'!${reference}`)
                if (relativeRange) return relativeRange
            }
            const matchingNames = definedNames.filter(item => item.getAttribute('name')?.toLowerCase() === name.toLowerCase())
            const localNames = sheetIndex === undefined ? [] : matchingNames.filter(item => {
                const localSheetId = item.getAttribute('localSheetId')
                return localSheetId !== null && /^\d+$/.test(localSheetId) && Number(localSheetId) === sheetIndex
            })
            const workbookNames = matchingNames.filter(item => item.getAttribute('localSheetId') === null)
            const candidates = localNames.length ? localNames : workbookNames
            if (candidates.length !== 1) return undefined
            const definition = candidates[0]!
            const definitionScope = definition.getAttribute('localSheetId')
            const definitionSheetIndex = definitionScope === null ? sheetIndex : /^\d+$/.test(definitionScope) ? Number(definitionScope) : undefined
            if (definitionScope !== null && (definitionSheetIndex === undefined || !worksheetNames[definitionSheetIndex])) return undefined
            const definitionKey = `${definitionScope ?? 'workbook'}:${name.toLowerCase()}`
            if (visited.has(definitionKey)) return undefined
            const nextVisited = new Set(visited)
            nextVisited.add(definitionKey)
            return chartRange(definition.textContent?.trim() || '', definitionSheetIndex, true, nextVisited, depth + 1)
        }
        for (const reference of [...descendants(chartRoot, 'numRef'), ...descendants(chartRoot, 'strRef')]) {
            const formula = child(reference, 'f')?.textContent?.trim()
            if (!formula) continue
            const range = chartRange(formula)
            if (!range) {
                warnings.add('Embedded chart workbook references must resolve to one-dimensional A1 ranges or defined names; unsupported references use saved chart caches.')
                continue
            }
            if (rangesByFormula.size >= MAX_CHART_REFERENCES) {
                warnings.add('Embedded chart workbook data and formats are limited to the first 512 references.')
                break
            }
            rangesByFormula.set(formula.toLowerCase(), range)
        }
        const refsBySheet = new Map<string, Set<string>>()
        let referencedCellCount = 0
        for (const range of rangesByFormula.values()) {
            const key = range.sheetName.toLowerCase()
            const refs = refsBySheet.get(key) || new Set<string>()
            for (const cell of range.cells) {
                if (!refs.has(cell)) {
                    if (referencedCellCount >= MAX_CHART_REFERENCED_CELLS) {
                        warnings.add('Embedded chart workbook data and formats are limited to 32,768 referenced cells.')
                        break
                    }
                    refs.add(cell)
                    referencedCellCount++
                }
            }
            refsBySheet.set(key, refs)
        }
        const cellsBySheet = new Map<string, Map<string, {numberFormat: string | undefined; value: string | null}>>()
        for (const sheet of children(firstDescendant(workbookXml.documentElement, 'sheets'), 'sheet')) {
            const sheetName = sheet.getAttribute('name') || ''
            const key = sheetName.toLowerCase()
            const requestedCells = refsBySheet.get(key)
            if (!requestedCells?.size) continue
            const relation = workbookRelationships.get(namespacedAttr(sheet, 'id') || '')
            if (!relation || relation.external || !relation.type.endsWith('/worksheet')) continue
            const worksheetPart = workbook.file(relation.target)
            if (!worksheetPart) continue
            const worksheetSize = (worksheetPart as JSZip.JSZipObject & {_data?: {uncompressedSize?: number}})._data?.uncompressedSize
            if (worksheetSize === undefined || worksheetSize > MAX_CHART_WORKSHEET_BYTES) {
                warnings.add('An embedded chart worksheet size is unavailable or exceeds the 8 MB reading limit; source-linked data and formats use saved chart caches.')
                continue
            }
            try {
                const worksheet = parseXml(await worksheetPart.async('string'), relation.target)
                const cells = new Map<string, {numberFormat: string | undefined; value: string | null}>()
                for (const cell of descendants(worksheet.documentElement, 'c')) {
                    const reference = cell.getAttribute('r')?.toUpperCase()
                    if (!reference || !requestedCells.has(reference)) continue
                    const rawValue = child(cell, 'v')?.textContent ?? null
                    const cellType = cell.getAttribute('t')
                    let value = rawValue?.trim() === '' ? null : rawValue
                    if (cellType === 's') {
                        const index = Number(rawValue)
                        value = Number.isSafeInteger(index) && index >= 0 ? sharedStrings[index] ?? null : null
                        if (value === null) warnings.add('An embedded chart workbook shared-string cell could not be resolved; affected category labels use point numbers.')
                    } else if (cellType === 'inlineStr') {
                        value = descendants(child(cell, 'is'), 't').map(text => text.textContent || '').join('')
                    } else if (cellType === 'b') {
                        value = rawValue === '1' ? 'TRUE' : rawValue === '0' ? 'FALSE' : null
                    } else if (cellType === 'e') {
                        value = null
                    }
                    cells.set(reference, {numberFormat: styleNumberFormat(cell.getAttribute('s')), value})
                }
                cellsBySheet.set(key, cells)
            } catch {
                warnings.add('An embedded chart worksheet could not be read; source-linked data and formats use saved chart caches.')
            }
        }
        const numberFormatsByFormula = new Map<string, Array<string | undefined>>()
        const valuesByFormula = new Map<string, Array<string | null>>()
        for (const [formula, range] of rangesByFormula) {
            const cells = cellsBySheet.get(range.sheetName.toLowerCase())
            numberFormatsByFormula.set(formula, range.cells.map(cell => cells?.get(cell)?.numberFormat))
            valuesByFormula.set(formula, range.cells.map(cell => cells?.get(cell)?.value ?? null))
        }
        return {
            ...(indexedColors ? {indexedColors: children(indexedColors, 'rgbColor').map(color => color.getAttribute('rgb') || '')} : {}),
            numberFormatsByFormula,
            valuesByFormula
        }
    } catch {
        warnings.add('An embedded chart workbook could not be read; indexed colors, source-linked values, and formats use their cache fallbacks.')
        return {numberFormatsByFormula: new Map(), valuesByFormula: new Map()}
    }
}

const EMBEDDED_FONT_VARIANTS = [
    {element: 'regular', weight: '400', style: 'normal'},
    {element: 'bold', weight: '700', style: 'normal'},
    {element: 'italic', weight: '400', style: 'italic'},
    {element: 'boldItalic', weight: '700', style: 'italic'}
] as const

function embeddedFontMime(data: Uint8Array): string | undefined {
    const signature = String.fromCharCode(...data.subarray(0, 4))
    if (signature === '\u0000\u0001\u0000\u0000' || signature === 'true' || signature === 'typ1') return 'font/ttf'
    if (signature === 'OTTO') return 'font/otf'
    if (signature === 'wOFF') return 'font/woff'
    if (signature === 'wOF2') return 'font/woff2'
    return undefined
}

async function parseEmbeddedFonts(
    presentation: XMLDocument,
    relationships: Map<string, PptxRelationship>,
    zip: JSZip,
    urls: string[],
    warnings: Set<string>
): Promise<PptxEmbeddedFontFace[]> {
    const list = firstDescendant(presentation.documentElement, 'embeddedFontLst')
    if (!list) return []

    const faces: PptxEmbeddedFontFace[] = []
    const seenFaces = new Set<string>()
    const urlsByTarget = new Map<string, string | null>()
    let totalBytes = 0
    let warnedSizeLimit = false
    let warnedFaceLimit = false

    for (const embeddedFont of children(list, 'embeddedFont')) {
        const family = child(embeddedFont, 'font')?.getAttribute('typeface')?.trim()
        if (!family) {
            warnings.add('An embedded font without a typeface name was skipped.')
            continue
        }

        for (const variant of EMBEDDED_FONT_VARIANTS) {
            const variantNode = child(embeddedFont, variant.element)
            if (!variantNode) continue
            if (faces.length >= MAX_EMBEDDED_FONT_FACES) {
                if (!warnedFaceLimit) warnings.add('Only the first 256 embedded font styles were read.')
                warnedFaceLimit = true
                continue
            }
            const relationId = namespacedAttr(variantNode, 'id')
            const relation = relationId ? relationships.get(relationId) : undefined
            if (!relation || relation.external || !relation.type.endsWith('/font')) {
                warnings.add(`The embedded ${variant.element} font for "${family}" has no valid internal font relationship and was skipped.`)
                continue
            }

            const faceKey = `${family.toLowerCase()}\u0000${variant.weight}\u0000${variant.style}`
            if (seenFaces.has(faceKey)) {
                warnings.add(`Duplicate embedded font style for "${family}" was ignored.`)
                continue
            }
            seenFaces.add(faceKey)

            let url = urlsByTarget.get(relation.target)
            if (url === undefined) {
                const entry = zip.file(relation.target)
                if (!entry) {
                    warnings.add(`The embedded font data for "${family}" is missing from the presentation.`)
                    urlsByTarget.set(relation.target, null)
                    continue
                }
                const size = (entry as JSZip.JSZipObject & {_data?: {uncompressedSize?: number}})._data?.uncompressedSize
                if (!Number.isFinite(size) || size! < 0) {
                    warnings.add(`The embedded font data for "${family}" has invalid ZIP size metadata and was skipped.`)
                    urlsByTarget.set(relation.target, null)
                    continue
                }
                if (totalBytes + size! > MAX_EMBEDDED_FONT_BYTES) {
                    if (!warnedSizeLimit) warnings.add('Embedded font data beyond the 64 MB browser limit was skipped.')
                    warnedSizeLimit = true
                    urlsByTarget.set(relation.target, null)
                    continue
                }
                const data = await entry.async('arraybuffer')
                if (totalBytes + data.byteLength > MAX_EMBEDDED_FONT_BYTES) {
                    if (!warnedSizeLimit) warnings.add('Embedded font data beyond the 64 MB browser limit was skipped.')
                    warnedSizeLimit = true
                    urlsByTarget.set(relation.target, null)
                    continue
                }
                totalBytes += data.byteLength
                const mime = embeddedFontMime(new Uint8Array(data))
                if (!mime) {
                    warnings.add(`The embedded font data for "${family}" is not a supported TrueType, OpenType, WOFF, or WOFF2 font.`)
                    urlsByTarget.set(relation.target, null)
                    continue
                }
                url = URL.createObjectURL(new Blob([data], {type: mime}))
                urls.push(url)
                urlsByTarget.set(relation.target, url)
            }

            if (url) faces.push({family, weight: variant.weight, style: variant.style, url})
        }
    }

    return faces
}

function parseHyperlink(node: Element | undefined, rels: Map<string, PptxRelationship> | undefined, warnings: Set<string>): PptxHyperlink | undefined {
    if (!node) return undefined
    const tooltip = node.getAttribute('tooltip') || undefined
    const action = (node.getAttribute('action') || '').trim().toLowerCase()
    const jump = /^ppaction:\/\/hlinkshowjump\?jump=(firstslide|lastslide|nextslide|previousslide|[1-9]\d*)$/.exec(action)?.[1]
    if (jump) {
        if (/^\d+$/.test(jump)) return {kind: 'showJump', jump: 'number', slideNumber: Number(jump), tooltip}
        return {kind: 'showJump', jump: jump as 'firstslide' | 'lastslide' | 'nextslide' | 'previousslide', tooltip}
    }

    const relationId = namespacedAttr(node, 'id')
    const relation = relationId ? rels?.get(relationId) : undefined
    if (action === 'ppaction://hlinksldjump' || (!action && relation?.type.endsWith('/slide'))) {
        if (relation && !relation.external && relation.type.endsWith('/slide')) return {kind: 'slide', targetPath: relation.target, tooltip}
        warnings.add('A slide hyperlink with an invalid relationship was ignored.')
        return undefined
    }
    if (action) {
        warnings.add('An unsupported presentation hyperlink action was ignored.')
        return undefined
    }
    if (!relation) {
        warnings.add('A hyperlink with a missing relationship was ignored.')
        return undefined
    }
    if (!relation.external || !relation.type.endsWith('/hyperlink')) {
        warnings.add('A non-external hyperlink relationship was ignored.')
        return undefined
    }
    try {
        const url = new URL(relation.target)
        if (['http:', 'https:', 'mailto:'].includes(url.protocol) && !url.username && !url.password) {
            return {kind: 'external', url: url.href, tooltip}
        }
    } catch {
        // Invalid or relative targets are not navigable URLs.
    }
    warnings.add('A hyperlink with an unsafe or unsupported URL was ignored.')
    return undefined
}

async function validateArchiveSize(zip: JSZip): Promise<void> {
    let total = 0
    let entryCount = 0
    for (const entry of Object.values(zip.files)) {
        if (++entryCount > MAX_ARCHIVE_ENTRIES) throw new Error('This presentation contains more than 20,000 package entries and exceeds the browser safety limit.')
        if (entry.dir) continue
        // ponytail: JSZip keeps expanded size private; replace this guard if it adds public metadata.
        const data: unknown = await (entry as JSZip.JSZipObject & {_data?: unknown})._data
        const size = typeof data === 'string' && data.length === 0
            ? 0
            : data !== null && typeof data === 'object' && 'uncompressedSize' in data
                ? data.uncompressedSize
                : undefined
        if (typeof size !== 'number' || !Number.isFinite(size) || size < 0) throw new Error('This presentation has incomplete or invalid ZIP size metadata and cannot be read safely.')
        total += size
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
    // HSL rotation can leave rounding noise at RGB endpoints; treat it as achromatic.
    if (delta <= Number.EPSILON * 8) return [0, 0, lightness]
    const saturation = delta / (1 - Math.abs(2 * lightness - 1))
    let hue = max === r ? ((g - b) / delta) % 6 : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4
    hue = (hue / 6 + 1) % 1
    return [hue, saturation, lightness]
}

function linearizeSrgbChannel(value: number): number {
    const sign = value < 0 ? -1 : 1
    const absolute = Math.abs(value)
    return sign * (absolute <= 0.04045 ? absolute / 12.92 : ((absolute + 0.055) / 1.055) ** 2.4)
}

function encodeSrgbChannel(value: number): number {
    const sign = value < 0 ? -1 : 1
    const absolute = Math.abs(value)
    return absolute > 0.0031308 ? sign * (1.055 * absolute ** (1 / 2.4) - 0.055) : value * 12.92
}

function readColor(node: Element | undefined, theme: Record<string, string>): string | undefined {
    if (!node) return undefined
    let hex = ''
    let hsl: [number, number, number] | undefined
    if (node.localName === 'srgbClr') hex = node.getAttribute('val') || ''
    else if (node.localName === 'sysClr') hex = node.getAttribute('lastClr') || '000000'
    else if (node.localName === 'schemeClr') hex = theme[node.getAttribute('val') || ''] || '#000000'
    else if (node.localName === 'hslClr') {
        const saturation = percentageValue(node.getAttribute('sat')) ?? 1
        const lightness = percentageValue(node.getAttribute('lum')) ?? 0.5
        hsl = [numberAttr(node, 'hue') / 21_600_000, Math.max(0, Math.min(1, saturation)), Math.max(0, Math.min(1, lightness))]
        const [r, g, b] = hslToRgb(...hsl)
        hex = [r, g, b].map(channel => Math.round(channel).toString(16).padStart(2, '0')).join('')
    } else if (node.localName === 'scrgbClr') {
        // scRGB channels are linear; encode them for the browser's sRGB colors.
        // https://www.w3.org/TR/css-color-4/#color-conversion-code
        hex = ['r', 'g', 'b'].map(name => {
            const linear = Math.max(0, Math.min(1, percentageValue(node.getAttribute(name)) ?? 0))
            const encoded = linear <= 0.0031308 ? linear * 12.92 : 1.055 * linear ** (1 / 2.4) - 0.055
            return Math.round(encoded * 255).toString(16).padStart(2, '0')
        }).join('')
    } else if (node.localName === 'prstClr') {
        const preset: Record<string, string> = {black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080', grey: '#808080'}
        hex = preset[node.getAttribute('val') || ''] || '#000000'
    }
    if (!hex) return undefined
    // Resolved style-reference colors may already carry alpha when used as phClr.
    const rgba = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*(?:,\s*([\d.]+)\s*)?\)$/i.exec(hex)
    if (!rgba && !hex.startsWith('#')) hex = `#${hex}`
    let r = rgba ? Number(rgba[1]) : parseInt(hex.slice(1, 3), 16)
    let g = rgba ? Number(rgba[2]) : parseInt(hex.slice(3, 5), 16)
    let b = rgba ? Number(rgba[3]) : parseInt(hex.slice(5, 7), 16)
    let alpha = rgba?.[4] === undefined ? 1 : Number(rgba[4])
    for (const mod of children(node)) {
        // Percentage strings are used by ISO OOXML; Office also writes thousandths of a percent.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/376a9585-53ee-4ab6-9a86-b58cafd200ba
        const kind = mod.localName
        if (kind === 'alphaInv') {
            alpha = 1 - Math.max(0, Math.min(1, alpha))
            continue
        }
        if (kind === 'inv') {
            ;[r, g, b] = [255 - r, 255 - g, 255 - b]
            hsl = undefined
            continue
        }
        if (kind === 'comp') {
            hsl ??= rgbToHsl(r, g, b)
            hsl[0] = ((hsl[0] + 0.5) % 1 + 1) % 1
            ;[r, g, b] = hslToRgb(...hsl)
            continue
        }
        if (kind === 'gray') {
            // PowerPoint's a:gray reference output weights encoded sRGB channels directly.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gray?view=openxml-3.0.1
            const intensity = 0.214 * r + 0.715 * g + 0.071 * b
            r = g = b = intensity
            hsl = undefined
            continue
        }
        if (kind === 'gamma' || kind === 'invGamma') {
            // DrawingML gamma transforms use the piecewise sRGB transfer curves.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gamma?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.inversegamma?view=openxml-3.0.1
            // https://www.w3.org/TR/css-color-4/#color-conversion-code
            const transformChannel = (channel: number) => (kind === 'gamma' ? encodeSrgbChannel(channel / 255) : linearizeSrgbChannel(channel / 255)) * 255
            ;[r, g, b] = [transformChannel(r), transformChannel(g), transformChannel(b)]
            hsl = undefined
            continue
        }
        const raw = mod.getAttribute('val')
        const value = kind === 'hue' || kind === 'hueOff'
            ? raw !== null && /^[+-]?\d+$/.test(raw.trim()) ? Number(raw) / 21_600_000 : undefined
            : percentageValue(raw)
        if (value === undefined || !Number.isFinite(value)) continue
        if (mod.localName === 'alpha') alpha = value
        else if (mod.localName === 'alphaMod') alpha = Math.max(0, Math.min(1, alpha * value))
        else if (mod.localName === 'alphaOff') alpha = Math.max(0, Math.min(1, alpha + value))
        else if (mod.localName === 'tint') {
            hsl = undefined
            // DrawingML mixes tint with white; PowerPoint applies it in linear-light sRGB.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tint?view=openxml-3.0.1
            r = encodeSrgbChannel(linearizeSrgbChannel(r / 255) * value + 1 - value) * 255
            g = encodeSrgbChannel(linearizeSrgbChannel(g / 255) * value + 1 - value) * 255
            b = encodeSrgbChannel(linearizeSrgbChannel(b / 255) * value + 1 - value) * 255
        } else if (mod.localName === 'shade') {
            hsl = undefined
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shade?view=openxml-3.0.1
            r = encodeSrgbChannel(linearizeSrgbChannel(r / 255) * value) * 255
            g = encodeSrgbChannel(linearizeSrgbChannel(g / 255) * value) * 255
            b = encodeSrgbChannel(linearizeSrgbChannel(b / 255) * value) * 255
        } else if (['hue', 'hueMod', 'hueOff', 'sat', 'satMod', 'satOff', 'lum', 'lumMod', 'lumOff'].includes(kind)) {
            // Preserve HSL between transforms, including hue while saturation is zero.
            // https://learn.microsoft.com/en-us/previous-versions/office/jj603569(v=office.15)
            hsl ??= rgbToHsl(r, g, b)
            const index = kind.startsWith('hue') ? 0 : kind.startsWith('sat') ? 1 : 2
            const transformed = kind.endsWith('Mod') ? hsl[index] * value : kind.endsWith('Off') ? hsl[index] + value : value
            hsl[index] = index === 0 ? ((transformed % 1) + 1) % 1 : Math.max(0, Math.min(1, transformed))
            ;[r, g, b] = hslToRgb(...hsl)
        } else if (['red', 'redMod', 'redOff', 'green', 'greenMod', 'greenOff', 'blue', 'blueMod', 'blueOff'].includes(kind)) {
            // Channel transforms leave the other components and alpha unchanged.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.redmodulation
            const channels: [number, number, number] = [r, g, b]
            const index = kind.startsWith('red') ? 0 : kind.startsWith('green') ? 1 : 2
            const transformed = kind.endsWith('Mod') ? channels[index] * value : kind.endsWith('Off') ? channels[index] + value * 255 : value * 255
            channels[index] = Math.max(0, Math.min(255, transformed))
            ;[r, g, b] = channels
            hsl = undefined
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
    if (node?.localName === 'hsl') {
        // PresentationML p:hsl values are relative HSL offsets, unlike DrawingML a:hslClr colors.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.hslcolor?view=openxml-3.0.1
        const rawHue = node.getAttribute('h')?.trim()
        const hue = rawHue && /^[+-]?(?:\d+\.?\d*|\.\d+)(?:deg)?$/i.test(rawHue)
            ? Number(rawHue.replace(/deg$/i, '')) / (rawHue.toLowerCase().endsWith('deg') ? 360 : 21_600_000)
            : Number.NaN
        const saturation = percentageValue(node.getAttribute('s'))
        const lightness = percentageValue(node.getAttribute('l'))
        if (colorSpace !== 'hsl' || !Number.isFinite(hue) || saturation === undefined || lightness === undefined) return undefined
        return [hue, saturation, lightness, 1]
    }
    if (node?.localName === 'rgb') {
        // PresentationML p:rgb stores relative channel offsets as fixed percentages.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.rgbcolor?view=openxml-3.0.1
        const red = percentageValue(node.getAttribute('r'))
        const green = percentageValue(node.getAttribute('g'))
        const blue = percentageValue(node.getAttribute('b'))
        if (colorSpace !== 'rgb' || red === undefined || green === undefined || blue === undefined) return undefined
        return [red * 255, green * 255, blue * 255, 1]
    }
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

type AnimationColorSpace = 'rgb' | 'hsl'
type AnimationColorValue = {colorSpace: AnimationColorSpace; channels: [number, number, number, number]}

// p:anim color strings use rgb(...), hsl(...), and RGB/ARGB hex forms.
// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7b427ccc-3a1f-418b-821f-d78f7aed51c5
function parseAnimationColorString(value: string | null | undefined): AnimationColorValue | undefined {
    const input = value?.trim()
    if (!input || input.length > 128) return undefined
    const hex = /^#([0-9a-f]{3}|[0-9a-f]{4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(input)?.[1]
    if (hex) {
        const byte = (value: string) => Number.parseInt(value, 16)
        const nibble = (value: string) => Number.parseInt(value, 16) * 16
        const channels = hex.length === 3
            ? [nibble(hex[0]!), nibble(hex[1]!), nibble(hex[2]!), 1]
            : hex.length === 4
              ? [nibble(hex[1]!), nibble(hex[2]!), nibble(hex[3]!), nibble(hex[0]!) / 255]
              : hex.length === 6
                ? [byte(hex.slice(0, 2)), byte(hex.slice(2, 4)), byte(hex.slice(4, 6)), 1]
                : [byte(hex.slice(2, 4)), byte(hex.slice(4, 6)), byte(hex.slice(6, 8)), byte(hex.slice(0, 2)) / 255]
        return {colorSpace: 'rgb', channels: channels as [number, number, number, number]}
    }

    const match = /^(rgb|hsl)\(\s*([^()]*)\s*\)$/i.exec(input)
    if (!match) return undefined
    const colorSpace = match[1]!.toLowerCase() as AnimationColorSpace
    const values = match[2]!.split(',').map(part => part.trim())
    if (values.length !== 3 && values.length !== 4) return undefined
    const parseChannel = (raw: string, maximum: number): number | undefined => {
        const isPercentage = raw.endsWith('%')
        const numericValue = (isPercentage ? raw.slice(0, -1) : raw).trim()
        if (!numericValue) return undefined
        const numeric = Number(numericValue)
        if (!Number.isFinite(numeric) || numeric < 0 || numeric > (isPercentage ? 100 : maximum)) return undefined
        return isPercentage ? numeric * maximum / 100 : numeric
    }
    const maximum = 255
    const redOrHue = parseChannel(values[0]!, maximum)
    const greenOrSaturation = parseChannel(values[1]!, maximum)
    const blueOrLightness = parseChannel(values[2]!, maximum)
    const alpha = values.length === 4 ? parseChannel(values[3]!, maximum) : maximum
    if (redOrHue === undefined || greenOrSaturation === undefined || blueOrLightness === undefined || alpha === undefined) return undefined
    const channels: [number, number, number, number] = colorSpace === 'hsl'
        ? [redOrHue / maximum, greenOrSaturation / maximum, blueOrLightness / maximum, alpha / maximum]
        : [redOrHue, greenOrSaturation, blueOrLightness, alpha / maximum]
    return {colorSpace, channels}
}

function animationColorValueInSpace(color: AnimationColorValue, colorSpace: AnimationColorSpace): [number, number, number, number] {
    if (color.colorSpace === colorSpace) return [...color.channels]
    const [first, second, third, alpha] = color.channels
    if (colorSpace === 'hsl') {
        const [hue, saturation, lightness] = rgbToHsl(first, second, third)
        return [hue, saturation, lightness, alpha]
    }
    const [red, green, blue] = hslToRgb(first, second, third)
    return [red, green, blue, alpha]
}

function animationColorStringChannels(value: string | null | undefined, colorSpace: AnimationColorSpace): [number, number, number, number] | undefined {
    const color = parseAnimationColorString(value)
    return color ? animationColorValueInSpace(color, colorSpace) : undefined
}

function parsedShadow(shadow: Element, theme: Record<string, string>, warnings: Set<string>, inset = false): PptxShadow {
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.innershadow?view=openxml-3.0.1
    const transformOverrides = [['sx', 100_000], ['sy', 100_000], ['kx', 0], ['ky', 0]] as const
    const hasTransformOverride = transformOverrides.some(([name, fallback]) => {
        const value = shadow.getAttribute(name)
        return value !== null && (!value.trim() || !Number.isFinite(Number(value)) || Number(value) !== fallback)
    })
    const alignment = shadow.getAttribute('algn')
    const rotation = shadow.getAttribute('rotWithShape')
    if (!inset && (hasTransformOverride || (alignment !== null && alignment !== 'b') || (rotation !== null && !['1', 'true'].includes(rotation)))) {
        warnings.add('An outer shadow uses scale, skew, alignment, or rotation settings that the CSS approximation does not preserve.')
    }
    return {
        color: readColor(children(shadow)[0], theme) || 'rgba(0, 0, 0, .35)',
        blur: numberAttr(shadow, 'blurRad'),
        distance: numberAttr(shadow, 'dist'),
        angle: numberAttr(shadow, 'dir') / 60_000,
        ...(inset ? {inset: true} : {})
    }
}

const CSS_INSET_SHADOW_GEOMETRIES = new Set(['rect', 'ellipse', 'roundRect', 'flowChartConnector', 'flowChartSummingJunction', 'flowChartOr', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartInternalStorage'])

function parsedInnerShadow(shadow: Element, theme: Record<string, string>, warnings: Set<string>, target: 'shape' | 'picture' | 'text', supported: boolean, supportsSvg = false): PptxShadow | undefined {
    if (!supported) {
        if (target === 'shape' && supportsSvg) {
            warnings.add('DrawingML inner shadows on custom geometries and SVG polygon presets use SVG filters and have not been calibrated against Office/WPS.')
            return parsedShadow(shadow, theme, warnings, true)
        }
        warnings.add(target === 'text'
            ? 'DrawingML inner shadows on text runs are not supported and were skipped.'
            : 'DrawingML inner shadows on preset geometries without SVG fill paths are not rendered.')
        return undefined
    }
    warnings.add(target === 'text'
        ? 'DrawingML inner shadows on text runs use SVG filters and have not been calibrated against Office/WPS.'
        : 'DrawingML inner shadow is approximated with CSS inset box-shadow and has not been calibrated against Office/WPS.')
    return parsedShadow(shadow, theme, warnings, true)
}

function effectShadows(parent: Element | undefined, theme: Record<string, string>, warnings: Set<string>, target: 'shape' | 'picture' | 'text', supportsInset: boolean, supportsSvg = false): PptxShadow[] {
    return children(parent).flatMap(shadow => {
        if (shadow.localName === 'outerShdw') return [parsedShadow(shadow, theme, warnings)]
        if (shadow.localName !== 'innerShdw') return []
        const parsed = parsedInnerShadow(shadow, theme, warnings, target, supportsInset, supportsSvg)
        return parsed ? [parsed] : []
    })
}

function parsedGlow(glow: Element, theme: Record<string, string>, warnings: Set<string>): PptxGlow | undefined {
    const radius = numberAttr(glow, 'rad', Number.NaN)
    const color = readColor(children(glow)[0], theme)
    if (!Number.isFinite(radius) || radius <= 0 || !color) {
        warnings.add('An invalid DrawingML glow effect was skipped.')
        return undefined
    }
    // DrawingML glow adds a blurred colored contour; CSS shadow rendering is pending Office/WPS calibration.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.glow?view=openxml-3.0.1
    warnings.add('DrawingML glow is approximated with CSS shadows and has not been calibrated against Office/WPS.')
    return {color, radius}
}

function glowEffects(parent: Element | undefined, theme: Record<string, string>, warnings: Set<string>): PptxGlow[] {
    const glows = parent?.localName === 'glow' ? [parent] : children(parent, 'glow')
    return glows.map(glow => parsedGlow(glow, theme, warnings)).filter((glow): glow is PptxGlow => Boolean(glow))
}

const FILL_KINDS = ['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill']

const patternLineStyles: Record<string, [number[], number, number]> = {
    horz: [[0], 1, 6], vert: [[90], 1, 6], cross: [[0, 90], 1, 6],
    ltHorz: [[0], 0.5, 6], ltVert: [[90], 0.5, 6], dkHorz: [[0], 2, 6], dkVert: [[90], 2, 6],
    narHorz: [[0], 1, 4], narVert: [[90], 1, 4], dashHorz: [[0], 1, 8], dashVert: [[90], 1, 8],
    dnDiag: [[45], 1, 6], upDiag: [[135], 1, 6], ltDnDiag: [[45], 0.5, 7], ltUpDiag: [[135], 0.5, 7],
    dkDnDiag: [[45], 2, 6], dkUpDiag: [[135], 2, 6], wdDnDiag: [[45], 3, 8], wdUpDiag: [[135], 3, 8],
    dashDnDiag: [[45], 1, 8], dashUpDiag: [[135], 1, 8], diagCross: [[45, 135], 1, 7],
    smGrid: [[0, 90], 1, 4], lgGrid: [[0, 90], 1, 10]
}

function patternFillColor(fill: Element, theme: Record<string, string>, warnings: Set<string>, allowCssPattern: boolean): string {
    const preset = fill.getAttribute('prst') || ''
    const foreground = readColor(children(child(fill, 'fgClr'))[0], theme) || 'transparent'
    const background = readColor(children(child(fill, 'bgClr'))[0], theme) || 'transparent'
    const fallback = background === 'transparent' ? foreground : background
    const density = /^pct(5|10|20|25|30|40|50|60|70|75|80|90)$/.exec(preset)
    const lines = patternLineStyles[preset]
    if (!density && !lines) {
        warnings.add(`The "${preset || 'unknown'}" DrawingML pattern fill is unsupported and was rendered as a solid color.`)
        return fallback
    }
    if (!allowCssPattern) {
        warnings.add(`The "${preset}" DrawingML pattern fill was rendered as a solid color in this context.`)
        return foreground === 'transparent' ? fallback : foreground
    }
    warnings.add('DrawingML pattern fills use approximate CSS rendering; text patterns repeat within each text run. SVG path fills use an embedded CSS viewport. Pattern scale may differ from Office/WPS.')
    if (density) {
        const spacing = 12
        const radius = spacing * Math.sqrt(Number(density[1]) / 100 / Math.PI)
        return `radial-gradient(circle ${radius.toFixed(2)}px at center, ${foreground} 99%, transparent 100%) 0 0 / ${spacing}px ${spacing}px, ${background}`
    }
    const [angles, width, spacing] = lines!
    const layers = angles.map(angle => `repeating-linear-gradient(${angle}deg, ${foreground} 0px, ${foreground} ${width}px, transparent ${width}px, transparent ${spacing}px)`)
    return `${layers.join(', ')}, ${background}`
}

function isPatternFillCss(value: string): boolean {
    return /repeating-linear-gradient\(/i.test(value) || /\) 0 0 \/ \d+(?:\.\d+)?px \d+(?:\.\d+)?px,/i.test(value)
}

function patternForegroundColor(value: string): string {
    const colors = [...value.matchAll(/transparent|#[\da-f]{3,8}\b|rgba?\([^)]*\)/gi)].map(match => match[0])
    return colors.find(color => !/^transparent$/i.test(color) && !/^rgba\([^,]+,[^,]+,[^,]+,\s*0(?:\.0+)?\s*\)$/i.test(color)) || 'transparent'
}

function fillColor(parent: Element | undefined, theme: Record<string, string>, warnings: Set<string>, groupFill?: string, allowCssPattern = false): string {
    const fill = parent && FILL_KINDS.includes(parent.localName) ? parent : children(parent).find(node => FILL_KINDS.includes(node.localName))
    // MS-OI29500: a line with omitted fill properties defaults to solidFill.
    if (!fill && parent?.localName === 'ln') {
        if (theme.phClr) return theme.phClr
        warnings.add('A line without fill properties or an inherited color uses a solid black fallback.')
        return '#000000'
    }
    if (!fill || fill.localName === 'noFill') return 'transparent'
    if (fill.localName === 'grpFill') {
        if (groupFill !== undefined) {
            if (isPatternFillCss(groupFill)) warnings.add('Inherited group patterns are applied separately to child shapes and may differ from PowerPoint shared group coordinates.')
            else if (groupFill.startsWith('linear-gradient(') || groupFill.startsWith('radial-gradient(')) {
                warnings.add('Inherited group gradients are applied separately to child shapes and may differ from PowerPoint shared group coordinates.')
            }
            return groupFill
        }
        warnings.add('The grp fill was simplified.')
        return 'transparent'
    }
    if (fill.localName === 'gradFill') {
        // OOXML gradient stops and clockwise angle:
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill.angle
        const stopNodes = children(child(fill, 'gsLst'), 'gs')
        const validStops = stopNodes.flatMap(stop => {
            const position = fixedPercentage(stop.getAttribute('pos'))
            if (position === undefined) {
                warnings.add('A gradient stop with an invalid or missing position was ignored.')
                return []
            }
            const color = readColor(children(stop)[0], theme) || 'transparent'
            return [{ color, position }]
        })
        if (validStops.length < 2) {
            warnings.add('A gradient with fewer than two stops was ignored.')
            return validStops[0]?.color || 'transparent'
        }
        const stops = validStops.map(stop => `${stop.color} ${stop.position * 100}%`)
        const path = child(fill, 'path')
        if (path) {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathgradientfill
            const pathType = path.getAttribute('path')
            if (pathType !== 'circle') warnings.add('A non-circular path gradient was approximated with a radial gradient.')
            const focusRect = child(path, 'fillToRect')
            let focusLayer = ''
            let focusX = 0.5, focusY = 0.5, focusRadius = 0
            if (focusRect) {
                // MS-OI29500 specifies a 50% base default for b; omitted l/r/t edges default to 0.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c0c046ec-a61d-405d-88fe-74d8487a37d7
                const edge = (name: string) => percentageValue(focusRect.getAttribute(name)) ?? (name === 'b' ? 0.5 : 0)
                const left = edge('l'), right = edge('r'), top = edge('t'), bottom = edge('b')
                const width = 1 - left - right, height = 1 - top - bottom
                if (width < 0 || height < 0) {
                    warnings.add('A path-gradient fillToRect has crossed focus edges and was ignored.')
                    focusX = 0.5
                    focusY = 0.5
                } else {
                    const ext = child(child(parent, 'xfrm'), 'ext')
                    const shapeAspect = numberAttr(ext, 'cx') / numberAttr(ext, 'cy', 1)
                    const aspect = Number.isFinite(shapeAspect) && shapeAspect > 0 ? shapeAspect : 1
                    if (pathType === 'circle') {
                        // Office interpolates between circles circumscribed around the shape and its focus rectangle.
                        // https://learn.microsoft.com/en-us/answers/questions/2247174/how-is-attribute-filltorect-evaluated-for-a-gradie
                        const focusCenterX = left + width / 2
                        const focusCenterY = top + height / 2
                        const focusRadiusY = Math.hypot(width * aspect, height) / 2
                        const focusRadiusX = focusRadiusY / aspect
                        const focusWidth = focusRadiusX * 2
                        const focusHeight = focusRadiusY * 2
                        const center = validStops.find(stop => stop.position === 0)?.color || validStops[0]!.color
                        focusLayer = `radial-gradient(ellipse closest-side, ${center} 99.999%, transparent 100%) left ${(focusCenterX - focusRadiusX) * 100}% top ${(focusCenterY - focusRadiusY) * 100}% / ${focusWidth * 100}% ${focusHeight * 100}% no-repeat, `
                    } else {
                        const focusLeft = left
                        const focusTop = top
                        const focusWidth = width
                        const focusHeight = height
                        focusX = Math.abs(1 - width) > Number.EPSILON * 2
                            ? (focusLeft - width * 0) / (1 - width) : focusLeft
                        focusY = Math.abs(1 - height) > Number.EPSILON * 2
                            ? (focusTop - height * 0) / (1 - height) : focusTop
                        const distance = (x: number, y: number) => Math.hypot((x - focusX) * aspect, y - focusY)
                        const outerRadius = Math.max(...[0, 1].flatMap(x => [0, 1].map(y => distance(x, y))))
                        const focusRadiusPixels = Math.max(...[focusLeft, focusLeft + focusWidth].flatMap(x => [focusTop, focusTop + focusHeight].map(y => distance(x, y))))
                        focusRadius = outerRadius ? Math.min(focusRadiusPixels / outerRadius, 1) : 0
                        if (focusWidth > 0 && focusHeight > 0) {
                            const center = validStops.find(stop => stop.position === 0)?.color || validStops[0]!.color
                            focusLayer = `linear-gradient(${center}, ${center}) left ${focusLeft * 100}% top ${focusTop * 100}% / ${focusWidth * 100}% ${focusHeight * 100}% no-repeat, `
                        }
                    }
                    if (focusLayer && pathType !== 'circle') warnings.add('A non-circular path-gradient focus rectangle is rendered over a radial-gradient approximation.')
                }
            }
            const radialStart = pathType === 'circle' ? 0 : focusRadius
            const radialStops = validStops.map(stop => `${stop.color} ${(radialStart + (1 - radialStart) * stop.position) * 100}%`)
            return `${focusLayer}radial-gradient(${pathType === 'circle' ? 'circle' : 'ellipse'} at ${focusX * 100}% ${focusY * 100}%, ${radialStops.join(', ')})`
        }
        const angle = numberAttr(child(fill, 'lin'), 'ang') / 60_000
        return `linear-gradient(${(90 + angle) % 360}deg, ${stops.join(', ')})`
    }
    if (fill.localName === 'pattFill') return patternFillColor(fill, theme, warnings, allowCssPattern)
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
        if (stops.length && pattern.length === stops.length * 2) {
            if (line?.getAttribute('cap') === 'rnd' && pattern.some((length, index) => index % 2 === 0 && length < 1)) {
                warnings.add('Custom round-cap dashes shorter than the pen width are approximated by SVG dots at the pen width.')
            }
            return pattern
        }
        warnings.add('A malformed custom line dash was approximated with a dash.')
        return 'dash'
    }
    return undefined
}

function lineCapStyle(line: Element | undefined, warnings: Set<string>): PptxElement['strokeCap'] {
    const value = line?.getAttribute('cap')
    // Office uses flat for omitted caps (MS-OI29500 2.1.1212).
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5564035f-88ff-4ab8-bbfb-d2998a83324d
    if (!value) return line ? 'butt' : undefined
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
    const guides = new Map<string, string>()
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

    // Guide declarations are evaluated in order; later declarations may reuse a prior guide name.
    for (const node of nodes.slice(0, 1024)) {
        const name = node.getAttribute('name'),
            formula = node.getAttribute('fmla')
        if (!name || !formula) continue
        guides.set(name, formula)
        values.delete(name)
        resolve(name)
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
// DrawingML arcTo locks its start angle to the current point; ellipse angles are polar angles before path scaling.
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
                const isValidAngle = (angle: number | undefined) => Number.isSafeInteger(angle) && angle! >= -2_147_483_648 && angle! <= 2_147_483_647
                if (!hasMove || !currentPoint || !Number.isFinite(radiusX) || !Number.isFinite(radiusY) || !(radiusX > 0) || !(radiusY > 0) || !isValidAngle(startAngleValue) || !isValidAngle(sweepAngleValue)) {
                    unsupported = true
                    break
                }
                const startAngle = startAngleValue! % 21_600_000
                const sweepAngle = sweepAngleValue!
                const ellipseOffset = (angle: number): [number, number] => {
                    const [x, y] = presetEllipsePoint(radiusXValue!, radiusYValue!, angle)
                    return [(x - 500) * radiusX / 500, (y - 500) * radiusY / 500]
                }
                const [startX, startY] = ellipseOffset(startAngle)
                const centerX = currentPoint[0] - startX
                const centerY = currentPoint[1] - startY
                const segmentCount = Math.ceil(Math.abs(sweepAngle) / 10_800_000)
                for (let segment = 1; segment <= segmentCount; segment++) {
                    const angle = startAngle + (sweepAngle * segment) / segmentCount
                    const [offsetX, offsetY] = ellipseOffset(angle)
                    currentPoint = [centerX + offsetX, centerY + offsetY]
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
        const modifier = fillMode === 'lighten' || fillMode === 'lightenLess' || fillMode === 'darken' || fillMode === 'darkenLess' ? fillMode : undefined
        if (!modifier && !['norm', 'none'].includes(fillMode)) warnings.add(`The custom geometry ${fillMode} path fill mode was simplified to normal fill.`)
        return [{d: d.join(' '), fill: fillMode !== 'none', fillMode: modifier, stroke: !['0', 'false'].includes((path.getAttribute('stroke') || 'true').toLowerCase())}]
    })
    if (unresolvedGuide) warnings.add('A custom geometry path references a missing, cyclic, or unsupported guide formula; dependent segments were omitted.')
    if (unsupported) warnings.add('Unsupported or malformed custom geometry path segments were omitted.')
    return paths.length ? paths : undefined
}

const textUnderlineStyles: Record<string, PptxRun['underlineStyle']> = {
    // Office does not support the DrawingML "words" type and loads it as one solid underline.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/d4795664-b08f-48b4-bb8a-a64e6de09778
    none: undefined, words: 'solid', sng: 'solid', dbl: 'double', heavy: undefined,
    dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashHeavy: 'dashed', dashLong: 'dashed', dashLongHeavy: 'dashed',
    dotDash: 'dashed', dotDashHeavy: 'dashed', dotDotDash: 'dashed', dotDotDashHeavy: 'dashed',
    wavy: 'wavy', wavyHeavy: 'wavy', wavyDbl: 'wavy',
}
const approximatedTextUnderlineValues = new Set(['words', 'heavy', 'dottedHeavy', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavyHeavy', 'wavyDbl'])

// DrawingML run properties permit every fill variant, including noFill and grpFill.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.runproperties?view=openxml-3.0.1
function textFillColor(fill: Element | undefined, theme: Record<string, string>, warnings: Set<string>, groupFill?: string, target: 'text' | 'underline' = 'text'): string | undefined {
    const label = target === 'text' ? 'text' : 'underline'
    const capitalizedLabel = target === 'text' ? 'Text' : 'Underline'
    if (!fill) return undefined
    if (fill.localName === 'solidFill') return readColor(children(fill)[0], theme)
    if (fill.localName === 'noFill') return 'transparent'
    if (fill.localName === 'grpFill') {
        if (groupFill && !groupFill.includes('gradient(')) return groupFill
        warnings.add(groupFill ? `${capitalizedLabel} group gradients use the inherited ${label} color because shared group coordinates are not supported.` : `A ${label} grpFill without a resolved group fill uses the inherited ${label} color.`)
        return undefined
    }
    if (fill.localName === 'pattFill') {
        if (target === 'underline') return patternFillColor(fill, theme, warnings, false)
        return readColor(children(child(fill, 'fgClr'))[0], theme) || readColor(children(child(fill, 'bgClr'))[0], theme)
    }
    if (fill.localName === 'gradFill') {
        if (target === 'underline') warnings.add(`Gradient ${label} fills are approximated with their first color stop.`)
        const firstStop = children(child(fill, 'gsLst'), 'gs')[0]
        return readColor(children(firstStop)[0], theme)
    }
    warnings.add(`The ${fill.localName.replace('Fill', '').toLowerCase()} ${label} fill is unsupported; inherited ${label} color is used.`)
    return undefined
}

function textFillPattern(fill: Element | undefined, theme: Record<string, string>, warnings: Set<string>): string | undefined {
    if (fill?.localName !== 'pattFill') return undefined
    const pattern = patternFillColor(fill, theme, warnings, true)
    return isPatternFillCss(pattern) ? pattern : undefined
}

function textFillGradient(fill: Element | undefined, theme: Record<string, string>, warnings: Set<string>): string | undefined {
    if (fill?.localName !== 'gradFill') return undefined
    const gradient = fillColor(fill, theme, warnings, undefined, true)
    if (!/(?:linear|radial)-gradient\(/i.test(gradient)) return undefined
    warnings.add('Text run gradients use CSS background clipping; gradient continuity across runs may differ from PowerPoint or WPS.')
    return gradient
}

// DrawingML a:highlight is a run-level color choice inside a:rPr.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.highlight?view=openxml-3.0.1
function textHighlightColor(style: Element | undefined, theme: Record<string, string>, warnings: Set<string>): string | undefined {
    const highlight = child(style, 'highlight')
    if (!highlight) return undefined
    const color = readColor(children(highlight)[0], theme)
    if (!color) {
        warnings.add('An invalid DrawingML text highlight color was ignored.')
        return undefined
    }
    warnings.add('DrawingML text highlights use browser inline backgrounds; line-fragment geometry may differ from PowerPoint or WPS.')
    return color
}

function textStyle(style: Element | undefined, theme: Record<string, string>, warnings: Set<string>, groupFill?: string): PptxTextStyle {
    const fill = children(style).find(node => FILL_KINDS.includes(node.localName))
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
        // Missing local font properties inherit; theme defaults belong at the bottom of the cascade.
        return value || undefined
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
    // uLn is an underline stroke; uLnTx follows the run outline. uFill controls its fill color.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underline?view=openxml-3.0.1
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underlinefollowstext?view=openxml-3.0.1
    // uFill sets underline paint; uFillTx follows this run's text color per DrawingML.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underlinefill?view=openxml-3.0.1
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underlinefilltext?view=openxml-3.0.1
    const underlineFill = child(style, 'uFill')
    const underlineFillFollowsText = Boolean(child(style, 'uFillTx'))
    const underlineLine = child(style, 'uLn')
    const underlineLineFollowsText = Boolean(child(style, 'uLnTx'))
    if (underlineFill && underlineFillFollowsText) warnings.add('A text run has both uFill and uFillTx; the explicit uFill is used.')
    if (underlineLine && underlineLineFollowsText) warnings.add('A text run has both uLn and uLnTx; the explicit uLn is used.')
    const underlineFillColor = underlineFill ? textFillColor(children(underlineFill)[0], theme, warnings, groupFill, 'underline') ?? 'currentColor' : undefined
    const underlineStrokeFill = children(underlineLine).find(node => FILL_KINDS.includes(node.localName))
    const underlineStrokeColor = underlineLine
        ? underlineStrokeFill ? textFillColor(underlineStrokeFill, theme, warnings, groupFill, 'underline') ?? 'currentColor' : undefined
        : underlineLineFollowsText ? outlineColor ?? 'currentColor' : undefined
    const underlineColor = underlineFillColor ?? (underlineFillFollowsText ? 'currentColor' : underlineStrokeColor)
    if (underlineFillColor && underlineStrokeColor && underlineFillColor !== underlineStrokeColor) warnings.add('Different DrawingML underline fill and stroke colors are simplified to the fill color.')
    const underlineStroke = underlineLine || (underlineLineFollowsText ? line : undefined)
    const underlineStrokeDash = lineDashStyle(underlineStroke, warnings)
    const hasUnderlineStrokeDash = Boolean(child(underlineStroke, 'prstDash') || child(underlineStroke, 'custDash'))
    let underlineStyle = validUnderline && rawUnderline !== undefined ? textUnderlineStyles[rawUnderline] : undefined
    if (hasUnderlineStrokeDash) {
        const strokeStyle = Array.isArray(underlineStrokeDash)
            ? 'dashed'
            : underlineStrokeDash === 'dot' || underlineStrokeDash === 'sysDot' ? 'dotted'
                : underlineStrokeDash === 'solid' ? 'solid' : 'dashed'
        if (Array.isArray(underlineStrokeDash)) warnings.add('A custom underline dash pattern is approximated with a CSS dashed underline.')
        if (underlineStyle && underlineStyle !== 'solid' && underlineStyle !== strokeStyle) warnings.add('Underline style and stroke dash are combined using the underline style.')
        else underlineStyle = strokeStyle
    }
    let underlineThicknessEmu: number | undefined
    const rawUnderlineWidth = underlineLine?.getAttribute('w')
    if (rawUnderlineWidth !== null && rawUnderlineWidth !== undefined) {
        const value = Number(rawUnderlineWidth)
        if (!/^\+?\d+$/.test(rawUnderlineWidth.trim()) || !Number.isSafeInteger(value) || value > 2_147_483_647) {
            warnings.add('A malformed DrawingML underline width was ignored.')
        } else if (value > 0) underlineThicknessEmu = value
    } else if (!underlineLine && underlineLineFollowsText && outlineWidth !== undefined && outlineWidth > 0) {
        underlineThicknessEmu = outlineWidth
    }
    if (underlineStroke && (['algn', 'cmpd', 'cap'].some(name => underlineStroke.hasAttribute(name)) || children(underlineStroke).some(node => ![...FILL_KINDS, 'prstDash', 'custDash', 'round', 'bevel', 'miter', 'extLst'].includes(node.localName)))) {
        warnings.add('Additional DrawingML underline stroke details are simplified.')
    }
    const strike = style?.getAttribute('strike')
    const validStrike = strike == null || ['noStrike', 'sngStrike', 'dblStrike'].includes(strike)
    if (!validStrike) warnings.add(`A text run with unsupported strike value "${strike}" was rendered without strikethrough.`)
    const effectContainer = child(style, 'effectLst') || child(style, 'effectDag')
    const dagEffects = effectContainer?.localName === 'effectDag' ? drawingEffectDag(effectContainer, theme, warnings, 'text', true) : undefined
    const shadows = dagEffects ? dagEffects.shadows || (dagEffects.shadow ? [dagEffects.shadow] : []) : effectShadows(effectContainer, theme, warnings, 'text', true)
    const glows = dagEffects ? dagEffects.glows || (dagEffects.glow ? [dagEffects.glow] : []) : glowEffects(effectContainer, theme, warnings)
    return {
        language: style?.getAttribute('lang')?.trim() || undefined,
        fontFamily: typeface('latin'),
        fontFamilyEastAsia: typeface('ea'),
        fontFamilyComplexScript: typeface('cs'),
        fontSizePt: style?.hasAttribute('sz') ? numberAttr(style, 'sz') / 100 : undefined,
        letterSpacingPt: validLetterSpacing ? letterSpacing / 100 : undefined,
        baselineOffsetEm: validBaseline && baselineValue !== undefined ? baselineValue / (baselineIsPercent ? 100 : 100_000) : undefined,
        bold: style?.hasAttribute('b') ? style.getAttribute('b') === '1' : undefined,
        italic: style?.hasAttribute('i') ? style.getAttribute('i') === '1' : undefined,
        underline: rawUnderline === undefined ? undefined : rawUnderline !== 'none',
        underlineStyle,
        underlineColor,
        underlineThicknessEmu,
        lineThrough: validStrike && strike != null ? strike !== 'noStrike' : undefined,
        lineThroughStyle: validStrike && strike != null ? strike === 'dblStrike' ? 'double' : 'solid' : undefined,
        color: textFillColor(fill, theme, warnings, groupFill),
        fillPattern: textFillPattern(fill, theme, warnings),
        fillGradient: textFillGradient(fill, theme, warnings),
        highlightColor: textHighlightColor(style, theme, warnings),
        outlineColor,
        outlineWidth,
        shadow: dagEffects ? dagEffects.shadow : shadows[0],
        shadows: shadows.length > 1 ? shadows : undefined,
        glow: glows[0],
        glows: glows.length > 1 ? glows : undefined
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

function resolveTextFonts(style: PptxTextStyle, text: string, theme: Record<string, string>, warnings: Set<string>): PptxTextStyle {
    let languageScript: string | undefined
    if (style.language) {
        try { languageScript = new Intl.Locale(style.language).maximize().script }
        catch { warnings.add('An invalid text language tag was ignored when resolving theme fonts.') }
    }
    // CT_SupplementalFont uses ISO script tags; DrawingML uses Hang for Korean.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.supplementalfont
    if (languageScript === 'Kore') languageScript = 'Hang'
    if (['Hira', 'Kana', 'Hrkt'].includes(languageScript || '')) languageScript = 'Jpan'
    const eastAsianScripts = ['Hans', 'Hant', 'Jpan', 'Hang']
    const inferredEastAsia = /[\p{Script=Hiragana}\p{Script=Katakana}]/u.test(text) ? 'Jpan'
        : /\p{Script=Hangul}/u.test(text) ? 'Hang' : /\p{Script=Han}/u.test(text) ? 'Hans' : undefined
    // ponytail: untagged Han defaults to Hans; the document's language is needed to distinguish Hant or Kanji-only Japanese.
    if (inferredEastAsia === 'Hans' && !eastAsianScripts.includes(languageScript || '')) warnings.add('Han text without an East Asian language uses the simplified Chinese theme-font mapping.')
    const inferredComplex = /\p{Script=Arabic}/u.test(text) ? 'Arab' : /\p{Script=Hebrew}/u.test(text) ? 'Hebr' : undefined
    const result = {...style}
    for (const key of ['fontFamily', 'fontFamilyEastAsia', 'fontFamilyComplexScript'] as const) {
        const value = style[key]
        const token = value && /^\+m([jn])-(lt|ea|cs)$/.exec(value)
        if (!token) continue
        const role = token[1] === 'j' ? 'major' : 'minor'
        const category = token[2] === 'lt' ? 'latin' : token[2]
        const script = category === 'ea' ? eastAsianScripts.includes(languageScript || '') ? languageScript : inferredEastAsia
            : category === 'cs' ? languageScript && !['Latn', 'Cyrl', 'Grek', ...eastAsianScripts].includes(languageScript) ? languageScript : inferredComplex
                : languageScript
        result[key] = theme[`font-${role}-${category}`] || (script ? theme[`font-${role}-script-${script}`] : undefined)
        if (!result[key] && ((category === 'ea' && inferredEastAsia) || (category === 'cs' && script))) warnings.add(`A ${role} ${category} theme font could not be resolved for this text; browser font fallback is used.`)
    }
    return result
}

function parseSpacing(properties: Element | undefined, name: string): PptxParagraph['lineSpacing'] {
    const spacing = child(properties, name)
    const points = child(spacing, 'spcPts')
    if (points) return {unit: 'emu', value: numberAttr(points, 'val') * 127}
    const percent = child(spacing, 'spcPct')
    if (percent) return {unit: 'ratio', value: numberAttr(percent, 'val') / 100_000}
    return undefined
}

function parseBullet(properties: Element | undefined, warnings: Set<string>, rels?: Map<string, PptxRelationship>): PptxParagraph['bullet'] {
    if (child(properties, 'buNone')) return undefined
    const percentSize = child(properties, 'buSzPct')
    const pointSize = child(properties, 'buSzPts')
    let size: NonNullable<PptxParagraph['bullet']>['size']
    if (percentSize && pointSize) warnings.add('A paragraph with both percentage and point bullet sizes was rendered with the default size.')
    else if (percentSize) {
        const value = numberAttr(percentSize, 'val', Number.NaN)
        if (Number.isSafeInteger(value) && value >= 25_000 && value <= 400_000) size = {unit: 'ratio', value: value / 100_000}
        else warnings.add('An invalid a:buSzPct value was ignored.')
    } else if (pointSize) {
        const value = numberAttr(pointSize, 'val', Number.NaN)
        if (Number.isSafeInteger(value) && value >= 100 && value <= 400_000) size = {unit: 'points', value: value / 100}
        else warnings.add('An invalid a:buSzPts value was ignored.')
    }
    const character = child(properties, 'buChar')
    if (character) return {kind: 'char', char: character.getAttribute('char') || '•', size}
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
            'alphaUcPeriod',
            'arabicDbPlain',
            'arabicDbPeriod',
            'ea1ChsPeriod',
            'ea1ChsPlain',
            'ea1ChtPeriod',
            'ea1ChtPlain',
            'thaiAlphaParenBoth',
            'thaiAlphaParenR',
            'thaiAlphaPeriod',
            'thaiNumParenBoth',
            'thaiNumParenR',
            'thaiNumPeriod',
            'hindiNumParenR',
            'hindiNumPeriod'
        ])
        if (!supported.has(type)) warnings.add(`The ${type} numbered bullet style is simplified to decimal numbering.`)
        return {kind: 'number', type, startAt: automatic.hasAttribute('startAt') ? numberAttr(automatic, 'startAt', 1) : undefined, size}
    }
    const picture = child(properties, 'buBlip')
    if (picture) {
        const blip = child(picture, 'blip')
        const relationId = namespacedAttr(blip, 'embed') || ''
        const relation = rels?.get(relationId)
        if (relationId && relation && !relation.external && relation.type.endsWith('/image')) return {kind: 'image', target: relation.target, size}
        warnings.add('A picture bullet with a missing, external, or non-image relationship was skipped.')
    }
    return undefined
}

function parseParagraphDefaults(properties: Element | undefined, theme: Record<string, string>, warnings: Set<string>, rels?: Map<string, PptxRelationship>, groupFill?: string): ParagraphDefaults {
    const values: ParagraphValues = {}
    const rawDirection = properties?.getAttribute('rtl')
    const direction = rawDirection?.trim()
    // a:pPr/@rtl controls the paragraph base direction; it remains independent of a:bodyPr/@rtlCol.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.righttoleft?view=openxml-3.0.1
    if (direction === 'true' || direction === '1') values.direction = 'rtl'
    else if (direction === 'false' || direction === '0') values.direction = 'ltr'
    else if (rawDirection !== undefined && rawDirection !== null) warnings.add('An invalid paragraph rtl value was ignored.')
    const alignment = properties?.getAttribute('algn')
    if (alignment) {
        const alignments: Record<string, PptxParagraph['align']> = {l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', thaiDist: 'justify'}
        if (alignments[alignment]) values.align = alignments[alignment]
    }
    if (properties?.hasAttribute('lvl')) values.level = numberAttr(properties, 'lvl')
    if (properties?.hasAttribute('marL')) values.marginLeft = numberAttr(properties, 'marL')
    if (properties?.hasAttribute('marR')) values.marginRight = numberAttr(properties, 'marR')
    if (properties?.hasAttribute('indent')) values.indent = numberAttr(properties, 'indent')
    if (properties?.hasAttribute('defTabSz')) {
        const rawDefaultTabSize = properties.getAttribute('defTabSz')?.trim()
        const defaultTabSize = rawDefaultTabSize && /^\d+$/.test(rawDefaultTabSize) ? Number(rawDefaultTabSize) : Number.NaN
        if (Number.isSafeInteger(defaultTabSize) && defaultTabSize > 0 && defaultTabSize <= 2_147_483_647) values.tabStopInterval = defaultTabSize
        else warnings.add('An invalid a:pPr/@defTabSz value was ignored.')
    }
    const spaceBefore = parseSpacing(properties, 'spcBef')
    const spaceAfter = parseSpacing(properties, 'spcAft')
    const lineSpacing = parseSpacing(properties, 'lnSpc')
    if (spaceBefore) values.spaceBefore = spaceBefore
    if (spaceAfter) values.spaceAfter = spaceAfter
    if (lineSpacing) values.lineSpacing = lineSpacing
    if (children(properties).some(item => ['buNone', 'buChar', 'buAutoNum', 'buBlip'].includes(item.localName))) {
        values.bullet = parseBullet(properties, warnings, rels)
    }
    const tabList = child(properties, 'tabLst')
    const tabStops = tabList ? children(tabList, 'tab').flatMap(tab => {
        const rawPosition = tab.getAttribute('pos')?.trim()
        const position = rawPosition && /^\d+$/.test(rawPosition) ? Number(rawPosition) : Number.NaN
        if (!Number.isSafeInteger(position) || position <= 0 || position > 2_147_483_647) {
            warnings.add('An invalid a:tab position was ignored.')
            return []
        }
        const rawAlignment = tab.getAttribute('algn')?.trim()
        const alignments = {l: 'left', ctr: 'center', r: 'right', dec: 'decimal'} as const
        const alignment = rawAlignment ? alignments[rawAlignment as keyof typeof alignments] : undefined
        if (rawAlignment && !alignment) warnings.add('An unsupported a:tab alignment was ignored.')
        return [{position, alignment}]
    }) : undefined
    return {values, text: textStyle(child(properties, 'defRPr') || child(properties, 'endParaRPr'), theme, warnings, groupFill), tabStops}
}

function mergeParagraphDefaults(...defaults: (ParagraphDefaults | undefined)[]): ParagraphDefaults {
    const values: ParagraphValues = {}
    let text: PptxTextStyle = {}
    let tabStops: ParagraphDefaults['tabStops']
    for (const current of defaults) {
        if (!current) continue
        Object.assign(values, current.values)
        text = mergeTextStyles(text, current.text)
        if (current.tabStops !== undefined) tabStops = current.tabStops
    }
    return {values, text, tabStops}
}

const DEFAULT_DRAWINGML_TAB_STOP_SIZE_EMU = 914_400

function cssTabStopIntervalSupported(values: ParagraphValues): boolean {
    return !(values.align && values.align !== 'left')
        && !(values.direction && values.direction !== 'ltr')
        && !(values.indent !== undefined && values.indent !== 0)
        && !values.bullet
}

function uniformLeftTabStopInterval(
    tabStops: ParagraphDefaults['tabStops'],
    values: ParagraphValues
): number | undefined {
    if (!tabStops || tabStops.length < 2 || !cssTabStopIntervalSupported(values)) return undefined
    const firstStop = tabStops[0]
    if (!firstStop) return undefined
    const interval = firstStop.position
    return tabStops.every((stop, index) => stop.alignment === 'left'
        && Number.isSafeInteger(interval * (index + 1))
        && stop.position === interval * (index + 1)) ? interval : undefined
}

function defaultTabStopInterval(values: ParagraphValues): number | undefined {
    if (!cssTabStopIntervalSupported(values)) return undefined
    return values.tabStopInterval ?? DEFAULT_DRAWINGML_TAB_STOP_SIZE_EMU
}

function mergeParagraphDefaultsByLevel(...maps: (ParagraphDefaultsByLevel | undefined)[]): ParagraphDefaultsByLevel {
    const result: ParagraphDefaultsByLevel = {}
    for (let level = 0; level < 9; level++) {
        const defaults = maps.map(map => map?.[level]).filter((item): item is ParagraphDefaults => Boolean(item))
        if (defaults.length) result[level] = mergeParagraphDefaults(...defaults)
    }
    return result
}

function paragraphDefaultsByLevel(styleList: Element | undefined, theme: Record<string, string>, warnings: Set<string>, rels?: Map<string, PptxRelationship>, groupFill?: string): ParagraphDefaultsByLevel {
    const fallback = parseParagraphDefaults(child(styleList, 'defPPr'), theme, warnings, rels, groupFill)
    const result: ParagraphDefaultsByLevel = {}
    for (let level = 0; level < 9; level++) {
        result[level] = mergeParagraphDefaults(fallback, parseParagraphDefaults(child(styleList, `lvl${level + 1}pPr`), theme, warnings, rels, groupFill))
    }
    return result
}

function placeholderParagraphDefaults(shape: Element, theme: Record<string, string>, warnings: Set<string>, rels?: Map<string, PptxRelationship>): ParagraphDefaultsByLevel {
    const textBody = child(shape, 'txBody')
    const result = paragraphDefaultsByLevel(child(textBody, 'lstStyle'), theme, warnings, rels)
    for (const paragraph of children(textBody, 'p')) {
        const properties = child(paragraph, 'pPr')
        const level = properties?.hasAttribute('lvl') ? numberAttr(properties, 'lvl') : 0
        const paragraphDefaults = parseParagraphDefaults(properties, theme, warnings, rels)
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
    inheritedParagraphStyles: ParagraphDefaultsByLevel = {},
    rels?: Map<string, PptxRelationship>,
    groupFill?: string
): {
    paragraphs: PptxParagraph[]
    margins: PptxElement['margins']
    verticalAlign: PptxElement['verticalAlign']
    textWrap?: PptxElement['textWrap']
    textHorizontalOverflow?: PptxElement['textHorizontalOverflow']
    textVerticalOverflow?: PptxElement['textVerticalOverflow']
    textOrientation?: PptxTextOrientation
    textAnchorCenter?: boolean
    textColumnCount?: number
    textColumnSpacing?: number
    textColumnsRightToLeft?: boolean
    textFontScale?: number
    textLineSpacingReduction?: number
    textAutoFit?: PptxElement['textAutoFit']
    textAutoFitDynamic?: boolean
} {
    const txBody = child(shape, 'txBody')
    if (!txBody) return {paragraphs: [], margins: {left: 0, right: 0, top: 0, bottom: 0}, verticalAlign: 'top'}
    const fontReference = child(child(shape, 'style'), 'fontRef')
    const fontIndex = fontReference?.getAttribute('idx')
    const referenceStyle: PptxTextStyle = {}
    // fontRef supplies theme font and color defaults below explicit text formatting.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontreference
    if (fontIndex === 'major' || fontIndex === 'minor' || fontIndex === 'none') {
        if (fontIndex !== 'none') {
            const role = fontIndex === 'major' ? 'mj' : 'mn'
            referenceStyle.fontFamily = `+${role}-lt`
            referenceStyle.fontFamilyEastAsia = `+${role}-ea`
            referenceStyle.fontFamilyComplexScript = `+${role}-cs`
        }
        referenceStyle.color = readColor(children(fontReference)[0], theme)
        if (referenceStyle.color) theme = {...theme, phClr: referenceStyle.color}
    } else if (fontReference) warnings.add('An invalid shape fontRef index was ignored; text fonts and colors inherit.')
    inheritedStyle = mergeTextStyles({
        fontFamily: '+mn-lt',
        fontFamilyEastAsia: '+mn-ea',
        fontFamilyComplexScript: '+mn-cs'
    }, referenceStyle, inheritedStyle)
    const body = child(txBody, 'bodyPr')
    const rawColumnCount = body?.getAttribute('numCol')
    const columnCount = rawColumnCount === null || rawColumnCount === undefined ? undefined : Number(rawColumnCount)
    const textColumnCount = columnCount !== undefined && Number.isSafeInteger(columnCount) && columnCount >= 1 && columnCount <= 16 ? columnCount : undefined
    if (columnCount !== undefined && textColumnCount === undefined) warnings.add('An invalid a:bodyPr numCol was ignored.')
    const rawColumnSpacing = body?.getAttribute('spcCol')
    const columnSpacing = rawColumnSpacing === null || rawColumnSpacing === undefined ? undefined : Number(rawColumnSpacing)
    const textColumnSpacing = columnSpacing !== undefined && Number.isSafeInteger(columnSpacing) && columnSpacing >= 0 && columnSpacing <= 2_147_483_647 ? columnSpacing : undefined
    if (columnSpacing !== undefined && textColumnSpacing === undefined) warnings.add('An invalid a:bodyPr spcCol was ignored.')
    const rawRightToLeftColumns = body?.getAttribute('rtlCol')
    const rightToLeftColumnsValue = rawRightToLeftColumns?.trim()
    const textColumnsRightToLeft = rightToLeftColumnsValue === undefined || rightToLeftColumnsValue === null
        ? undefined
        : ['true', '1'].includes(rightToLeftColumnsValue)
            ? true
            : ['false', '0'].includes(rightToLeftColumnsValue)
                ? false
                : undefined
    if (rawRightToLeftColumns !== undefined && rawRightToLeftColumns !== null && textColumnsRightToLeft === undefined) warnings.add('An invalid a:bodyPr rtlCol value was ignored.')
    const normalAutoFit = child(body, 'normAutofit')
    const shapeAutoFit = child(body, 'spAutoFit')
    const rawFontScale = normalAutoFit?.getAttribute('fontScale') ?? null
    const rawLineSpaceReduction = normalAutoFit?.getAttribute('lnSpcReduction') ?? null
    const parsedFontScale = fixedPercentage(rawFontScale)
    const parsedLineSpaceReduction = fixedPercentage(rawLineSpaceReduction)
    if (shapeAutoFit) warnings.add('a:spAutoFit shape resizing uses browser font metrics; PowerPoint may produce different dimensions.')
    if (rawFontScale !== null && parsedFontScale === undefined) warnings.add('An invalid a:normAutofit fontScale was ignored.')
    if (rawLineSpaceReduction !== null && parsedLineSpaceReduction === undefined) warnings.add('An invalid a:normAutofit lnSpcReduction was ignored.')
    const listParagraphStyles = paragraphDefaultsByLevel(child(txBody, 'lstStyle'), theme, warnings, rels, groupFill)
    const paragraphs: PptxParagraph[] = children(txBody, 'p').map(paragraph => {
        const props = child(paragraph, 'pPr')
        const level = props?.hasAttribute('lvl') ? numberAttr(props, 'lvl') : 0
        const resolved = mergeParagraphDefaults(inheritedParagraphStyles[level], listParagraphStyles[level], parseParagraphDefaults(props, theme, warnings, rels, groupFill))
        const paragraphStyle = mergeTextStyles(inheritedStyle, resolved.text)
        const runs: PptxRun[] = []
        for (const run of children(paragraph)) {
            if (run.localName === 'br') {
                // DrawingML a:br is a forced line break between runs in one paragraph.
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.break
                const styleNode = child(run, 'rPr')
                const style = mergeTextStyles(paragraphStyle, textStyle(styleNode, theme, warnings, groupFill))
                const hyperlink = parseHyperlink(child(styleNode, 'hlinkClick'), rels, warnings)
                runs.push({text: '\n', ...resolveTextFonts(style, '\n', theme, warnings), ...(hyperlink ? {hyperlink} : {})})
                continue
            }
            if (!['r', 'fld'].includes(run.localName)) continue
            const textNode = child(run, 't')
            if (!textNode) continue
            const style = child(run, 'rPr') || child(run, 'endParaRPr')
            const runStyle = mergeTextStyles(paragraphStyle, textStyle(style, theme, warnings, groupFill))
            const hyperlink = parseHyperlink(child(style, 'hlinkClick'), rels, warnings)
            const text = textNode.textContent || ''
            runs.push({text, ...resolveTextFonts(runStyle, text, theme, warnings), ...(hyperlink ? {hyperlink} : {})})
        }
        if (!runs.length && paragraph.textContent?.trim()) {
            const text = paragraph.textContent.trim()
            runs.push({text, ...resolveTextFonts(paragraphStyle, text, theme, warnings)})
        }
        const hasTabCharacters = runs.some(run => run.text.includes('\t'))
        const tabStopInterval = hasTabCharacters
            ? uniformLeftTabStopInterval(resolved.tabStops, resolved.values)
                ?? (!resolved.tabStops?.length ? defaultTabStopInterval(resolved.values) : undefined)
            : undefined
        if (hasTabCharacters && tabStopInterval === undefined) {
            warnings.add('Non-uniform, non-left-aligned, RTL, indented, or bulleted text tab stops use browser default spacing.')
        }
        return {
            ...resolved.values,
            align: resolved.values.align || 'left',
            tabStopInterval,
            runs,
            level: props?.hasAttribute('lvl') ? level : resolved.values.level
        }
    })
    const margin = (key: string, fallback: number) => numberAttr(body, key, fallback)
    const anchor = body?.getAttribute('anchor')
    const wrap = body?.getAttribute('wrap')
    // DrawingML defaults both overflow axes to visible; clip constrains text to the body bounds.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.horizontaloverflow?view=openxml-3.0.1
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.verticaloverflow?view=openxml-3.0.1
    const horizontalOverflow = body?.getAttribute('horzOverflow')
    const verticalOverflow = body?.getAttribute('vertOverflow')
    const textHorizontalOverflow = horizontalOverflow === 'clip' || horizontalOverflow === 'overflow' ? horizontalOverflow : undefined
    const textVerticalOverflow = verticalOverflow === 'clip' || verticalOverflow === 'overflow' || verticalOverflow === 'ellipsis' ? verticalOverflow : undefined
    if (horizontalOverflow && !textHorizontalOverflow) warnings.add(`The unsupported text horizontal overflow mode "${horizontalOverflow}" was ignored.`)
    if (verticalOverflow && !textVerticalOverflow) warnings.add(`The unsupported text vertical overflow mode "${verticalOverflow}" was ignored.`)
    const rawTextOrientation = body?.getAttribute('vert')
    const validTextOrientations: PptxTextOrientation[] = ['horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl']
    const rawAnchorCenter = body?.getAttribute('anchorCtr')?.trim()
    const textAnchorCenter = rawAnchorCenter === null || rawAnchorCenter === undefined
        ? undefined
        : ['1', 'true'].includes(rawAnchorCenter)
            ? true
            : ['0', 'false'].includes(rawAnchorCenter)
                ? false
                : undefined
    if (rawAnchorCenter !== null && rawAnchorCenter !== undefined && textAnchorCenter === undefined) warnings.add('An invalid a:bodyPr anchorCtr value was ignored.')
    if (textVerticalOverflow === 'ellipsis' && ((textColumnCount ?? 1) > 1 || rawTextOrientation && rawTextOrientation !== 'horz')) {
        warnings.add('a:bodyPr vertOverflow="ellipsis" is clipped for vertical or multi-column text.')
    }
    if (rawTextOrientation && !validTextOrientations.includes(rawTextOrientation as PptxTextOrientation)) {
        warnings.add('The unsupported text vertical mode "' + rawTextOrientation + '" was ignored.')
    } else if (rawTextOrientation === 'wordArtVert' || rawTextOrientation === 'mongolianVert') {
        warnings.add(`The text vertical mode "${rawTextOrientation}" uses CSS writing modes; glyph shaping may differ from Office.`)
    }
    if (rawTextOrientation === 'wordArtVertRtl' && (textColumnCount ?? 1) > 1 && textColumnsRightToLeft !== true) {
        warnings.add('wordArtVertRtl requires right-to-left column flow; rtlCol="0" or an omitted rtlCol cannot request left-to-right columns, so its vertical flow is preserved.')
    }
    if (wrap && !['square', 'none'].includes(wrap)) warnings.add(`The unsupported text wrapping mode "${wrap}" was ignored.`)
    return {
        paragraphs,
        margins: {left: margin('lIns', 91_440), right: margin('rIns', 91_440), top: margin('tIns', 45_720), bottom: margin('bIns', 45_720)},
        verticalAlign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
        textWrap: wrap === 'square' || wrap === 'none' ? wrap : undefined,
        textHorizontalOverflow,
        textVerticalOverflow,
        textOrientation: rawTextOrientation && validTextOrientations.includes(rawTextOrientation as PptxTextOrientation) ? (rawTextOrientation as PptxTextOrientation) : undefined,
        textAnchorCenter,
        textColumnCount,
        textColumnSpacing,
        textColumnsRightToLeft,
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

function parseSpeakerNotes(xml: XMLDocument | undefined, master: XMLDocument | undefined, theme: Record<string, string>, warnings: Set<string>, rels: Map<string, PptxRelationship>, masterRels: Map<string, PptxRelationship>): Pick<PptxSlide, 'speakerNotes' | 'speakerNoteParagraphs'> {
    const tree = firstDescendant(xml?.documentElement, 'spTree')
    const bodies = children(tree, 'sp').filter(shape => parsePlaceholderIdentity(shape)?.type === 'body')
    const masterBodies = children(firstDescendant(master?.documentElement, 'spTree'), 'sp').filter(shape => parsePlaceholderIdentity(shape)?.type === 'body')
    // Notes have their own master text defaults and color map, independent of the slide master.
    // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides
    const defaults = paragraphDefaultsByLevel(child(master?.documentElement, 'notesStyle'), theme, warnings, masterRels)
    const paragraphs = bodies.flatMap(shape => {
        const identity = parsePlaceholderIdentity(shape)
        const placeholder = masterBodies.find(body => parsePlaceholderIdentity(body)?.index === identity?.index) || masterBodies[0]
        const inherited = mergeParagraphDefaultsByLevel(defaults, placeholder ? placeholderParagraphDefaults(placeholder, theme, warnings, masterRels) : undefined)
        return parseText(shape, theme, warnings, {}, inherited, rels).paragraphs
    })
    const text = paragraphs
        .map(paragraph => paragraph.runs.map(run => run.text).join(''))
        .join('\n')
        .trim()
    return text ? {speakerNotes: text, speakerNoteParagraphs: paragraphs} : {}
}

function placeholderStyleCategory(type: string): 'titleStyle' | 'bodyStyle' | 'otherStyle' {
    if (['title', 'ctrTitle'].includes(type)) return 'titleStyle'
    if (['dt', 'ftr', 'hdr', 'sldNum'].includes(type)) return 'otherStyle'
    return 'bodyStyle'
}

// Slide-master txStyles carries separate title, body, and other text defaults.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.textstyles?view=openxml-3.0.1
function parseMasterTextStyles(xml: XMLDocument | undefined, theme: Record<string, string>, warnings: Set<string>, rels?: Map<string, PptxRelationship>): Record<string, MasterTextStyle> {
    const styles = child(xml?.documentElement, 'txStyles')
    const result: Record<string, MasterTextStyle> = {}
    for (const name of ['titleStyle', 'bodyStyle', 'otherStyle']) {
        const style = child(styles, name)
        result[name] = {
            style: textStyle(child(style, 'defRPr'), theme, warnings),
            paragraphs: paragraphDefaultsByLevel(style, theme, warnings, rels)
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

function inheritPlaceholderProperties(element: PptxElement, node: Element, fallback: ParsedPlaceholder | undefined, warnings: Set<string>): PptxElement {
    if (!fallback) return element
    const fallbackElement = fallback.element
    const properties = child(node, 'spPr')
    const transform = child(properties, 'xfrm')
    const offset = child(transform, 'off')
    const extent = child(transform, 'ext')
    const fillNames = ['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill']
    const hasFill = children(properties).some(item => fillNames.includes(item.localName)) || !!child(child(node, 'style'), 'fillRef')
    const line = child(properties, 'ln')
    const hasLineReference = !!child(child(node, 'style'), 'lnRef')
    const hasLineFill = hasLineReference || children(line).some(item => fillNames.includes(item.localName))
    const hasEffectProperties = children(properties).some(item => ['effectLst', 'effectDag'].includes(item.localName)) || !!child(child(node, 'style'), 'effectRef')
    const bodyPr = child(child(node, 'txBody'), 'bodyPr')
    const hasTextAutoFit = Boolean(child(bodyPr, 'normAutofit') || child(bodyPr, 'spAutoFit') || child(bodyPr, 'noAutofit'))
    const fallbackMargins = fallbackElement.margins
    const margin = (name: 'lIns' | 'rIns' | 'tIns' | 'bIns', fallbackValue: number) => {
        const current = name === 'lIns' ? element.margins.left : name === 'rIns' ? element.margins.right : name === 'tIns' ? element.margins.top : element.margins.bottom
        return bodyPr?.hasAttribute(name) ? current : fallbackValue
    }
    const anchor = bodyPr?.getAttribute('anchor')
    const currentGeometry = child(properties, 'prstGeom') || child(properties, 'custGeom')
    const textOrientation = bodyPr?.hasAttribute('vert') ? element.textOrientation : fallbackElement.textOrientation
    const textAnchorCenter = bodyPr?.hasAttribute('anchorCtr') ? element.textAnchorCenter : fallbackElement.textAnchorCenter
    const textColumnCount = bodyPr?.hasAttribute('numCol') ? element.textColumnCount : fallbackElement.textColumnCount
    const textColumnsRightToLeft = bodyPr?.hasAttribute('rtlCol') ? element.textColumnsRightToLeft : fallbackElement.textColumnsRightToLeft
    if (textOrientation === 'wordArtVertRtl' && (textColumnCount ?? 1) > 1 && textColumnsRightToLeft !== true) {
        warnings.add('wordArtVertRtl requires right-to-left column flow; rtlCol="0" or an omitted rtlCol cannot request left-to-right columns, so its vertical flow is preserved.')
    }
    return {
        ...element,
        x: offset?.hasAttribute('x') ? element.x : fallbackElement.x,
        y: offset?.hasAttribute('y') ? element.y : fallbackElement.y,
        width: extent?.hasAttribute('cx') ? element.width : fallbackElement.width,
        height: extent?.hasAttribute('cy') ? element.height : fallbackElement.height,
        rotation: transform?.hasAttribute('rot') ? element.rotation : fallbackElement.rotation,
        flipH: transform?.hasAttribute('flipH') ? element.flipH : fallbackElement.flipH,
        flipV: transform?.hasAttribute('flipV') ? element.flipV : fallbackElement.flipV,
        kind: currentGeometry ? element.kind : fallbackElement.kind,
        geometry: currentGeometry ? element.geometry : fallbackElement.geometry,
        linePath: currentGeometry ? element.linePath : fallbackElement.linePath,
        presetAdjustment: currentGeometry ? element.presetAdjustment : fallbackElement.presetAdjustment,
        presetAdjustments: currentGeometry ? element.presetAdjustments : fallbackElement.presetAdjustments,
        customPaths: currentGeometry ? element.customPaths : fallbackElement.customPaths,
        customTextRect: currentGeometry ? element.customTextRect : fallbackElement.customTextRect,
        fill: hasFill ? element.fill : fallbackElement.fill,
        stroke: hasLineFill ? element.stroke : fallbackElement.stroke,
        strokeWidth: hasLineReference || line?.hasAttribute('w') ? element.strokeWidth : fallbackElement.strokeWidth,
        strokeDash: hasLineReference || child(line, 'prstDash') || child(line, 'custDash') ? element.strokeDash : fallbackElement.strokeDash,
        strokeCap: hasLineReference || line?.hasAttribute('cap') ? element.strokeCap : fallbackElement.strokeCap,
        lineHeadEnd: hasLineReference || child(line, 'headEnd') ? element.lineHeadEnd : fallbackElement.lineHeadEnd,
        lineTailEnd: hasLineReference || child(line, 'tailEnd') ? element.lineTailEnd : fallbackElement.lineTailEnd,
        // A slide placeholder inherits visual properties from its matching layout/master placeholder.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/b3e45aed-2196-42bd-8b88-c98dba85ffe4
        shadow: hasEffectProperties ? element.shadow : fallbackElement.shadow,
        shadows: hasEffectProperties ? element.shadows : fallbackElement.shadows,
        glow: hasEffectProperties ? element.glow : fallbackElement.glow,
        glows: hasEffectProperties ? element.glows : fallbackElement.glows,
        margins: {
            left: margin('lIns', fallbackMargins.left),
            right: margin('rIns', fallbackMargins.right),
            top: margin('tIns', fallbackMargins.top),
            bottom: margin('bIns', fallbackMargins.bottom)
        },
        verticalAlign: anchor ? element.verticalAlign : fallbackElement.verticalAlign,
        textWrap: ['square', 'none'].includes(bodyPr?.getAttribute('wrap') || '') ? element.textWrap : fallbackElement.textWrap,
        textHorizontalOverflow: bodyPr?.hasAttribute('horzOverflow') ? element.textHorizontalOverflow : fallbackElement.textHorizontalOverflow,
        textVerticalOverflow: bodyPr?.hasAttribute('vertOverflow') ? element.textVerticalOverflow : fallbackElement.textVerticalOverflow,
        textOrientation,
        textAnchorCenter,
        textColumnCount,
        textColumnSpacing: bodyPr?.hasAttribute('spcCol') ? element.textColumnSpacing : fallbackElement.textColumnSpacing,
        textColumnsRightToLeft,
        textFontScale: hasTextAutoFit ? element.textFontScale : fallbackElement.textFontScale,
        textLineSpacingReduction: hasTextAutoFit ? element.textLineSpacingReduction : fallbackElement.textLineSpacingReduction,
        textAutoFit: hasTextAutoFit ? element.textAutoFit : fallbackElement.textAutoFit,
        textAutoFitDynamic: hasTextAutoFit ? element.textAutoFitDynamic : fallbackElement.textAutoFitDynamic
    }
}

function parsePlaceholderCandidates(xml: XMLDocument | undefined, theme: Record<string, string>, warnings: Set<string>, masterStyles: Record<string, MasterTextStyle>, fallbacks: ParsedPlaceholder[] = [], warnPlaceholders = false, rels?: Map<string, PptxRelationship>, fillStyles?: ParsedTheme): ParsedPlaceholder[] {
    const tree = xml && firstDescendant(xml.documentElement, 'spTree')
    const candidates: ParsedPlaceholder[] = []
    for (const node of children(tree)) {
        if (node.localName !== 'sp') continue
        const identity = parsePlaceholderIdentity(node)
        if (!identity) continue
        const fallback = matchPlaceholder(identity, fallbacks)
        const masterStyle = masterStyles[placeholderStyleCategory(identity.type)]
        const paragraphStyles = mergeParagraphDefaultsByLevel(masterStyle?.paragraphs, fallback?.paragraphStyles, placeholderParagraphDefaults(node, theme, warnings, rels))
        let element = parseShape(node, theme, warnings, masterStyle?.style || {}, paragraphStyles, warnPlaceholders, rels, fillStyles)
        element = inheritPlaceholderProperties(element, node, fallback, warnings)
        candidates.push({node, type: identity.type, index: identity.index, element, paragraphStyles})
    }
    return candidates
}

function referencedShapeFill(shape: Element, theme: Record<string, string>, warnings: Set<string>, styles?: ParsedTheme, allowCssPattern = false): string {
    const reference = child(child(shape, 'style'), 'fillRef')
    if (!reference) return 'transparent'
    const raw = reference.getAttribute('idx') || ''
    const index = /^\d+$/.test(raw) ? Number(raw) : Number.NaN
    if (index === 0 || index === 1000) return 'transparent'
    const style = Number.isSafeInteger(index) && index > 0 ? index < 1000 ? styles?.fillStyles[index - 1] : styles?.backgroundFillStyles[index - 1001] : undefined
    if (!style) {
        warnings.add('An invalid or missing shape fillRef style was rendered without fill.')
        return 'transparent'
    }
    const color = readColor(children(reference)[0], theme)
    return fillColor(style, color ? {...theme, phClr: color} : theme, warnings, undefined, allowCssPattern)
}

function referencedShapeLine(shape: Element, theme: Record<string, string>, warnings: Set<string>, styles?: ParsedTheme): {line?: Element; theme: Record<string, string>} {
    const local = child(child(shape, 'spPr'), 'ln')
    const reference = child(child(shape, 'style'), 'lnRef')
    if (!reference) return {line: local, theme}
    const raw = reference.getAttribute('idx') || ''
    const index = /^\d+$/.test(raw) ? Number(raw) : Number.NaN
    const base = Number.isSafeInteger(index) && index > 0 ? styles?.lineStyles[index - 1] : undefined
    if (!base && index !== 0) warnings.add('An invalid or missing shape lnRef style was rendered using local line properties.')
    if (!base) return {line: local, theme}
    const line = base.cloneNode(true) as Element
    for (const attribute of Array.from(local?.attributes || [])) line.setAttribute(attribute.name, attribute.value)
    const groups = [['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'], ['prstDash', 'custDash'], ['round', 'bevel', 'miter']]
    for (const property of children(local)) {
        const names = groups.find(group => group.includes(property.localName)) || [property.localName]
        for (const inherited of children(line)) if (names.includes(inherited.localName)) line.removeChild(inherited)
        line.appendChild(property.cloneNode(true))
    }
    const color = readColor(children(reference)[0], theme)
    return {line, theme: color ? {...theme, phClr: color} : theme}
}

function singleAlphaInverseEffect(container: Element): Element | undefined {
    let current = container
    while (true) {
        const entries = children(current)
        if (entries.length !== 1) return undefined
        const entry = entries[0]
        if (!entry) return undefined
        if (entry.localName === 'alphaInv') return entry
        if (entry.localName !== 'cont' || !['sib', 'tree'].includes(entry.getAttribute('type') || 'sib')) return undefined
        current = entry
    }
}

function treeAlphaInverseEffectSequence(container: Element): Element[] | undefined {
    const type = container.getAttribute('type') || 'sib'
    const entries = children(container)
    if (!['sib', 'tree'].includes(type) || (type === 'sib' && entries.length > 1)) return undefined
    const effects: Element[] = []
    for (const entry of entries) {
        if (entry.localName === 'alphaInv') effects.push(entry)
        else if (entry.localName === 'cont') {
            const nested = treeAlphaInverseEffectSequence(entry)
            if (!nested) return undefined
            effects.push(...nested)
        } else return undefined
    }
    return effects.length ? effects : undefined
}

function drawingEffectDag(effects: Element, theme: Record<string, string>, warnings: Set<string>, target: 'shape' | 'picture' | 'text', supportsInset: boolean, supportsSvg = false): {shadow?: PptxShadow; shadows?: PptxShadow[]; glow?: PptxGlow; glows?: PptxGlow[]; alphaInverse?: PptxAlphaInverse} {
    const type = effects.getAttribute('type') || 'sib'
    const entries = children(effects)
    if (type !== 'sib' && type !== 'tree') {
        warnings.add(`The ${target} effect DAG container type "${type}" is not supported.`)
        return {}
    }
    const multiTreeAlphaInverseNodes = target === 'picture' && type === 'tree' && entries.length > 1
        ? treeAlphaInverseEffectSequence(effects)
        : undefined
    if (type === 'tree' && entries.length > 1 && !multiTreeAlphaInverseNodes) {
        // ECMA-376: tree containers apply each effect to the previous effect's result.
        warnings.add(`Multi-effect tree-ordered ${target} effect DAGs are not rendered.`)
        return {}
    }

    const shadows: PptxShadow[] = []
    const glows: PptxGlow[] = []
    const alphaInverseEntry = target === 'picture' && entries.length === 1 ? entries[0] : undefined
    const alphaInverseNode = alphaInverseEntry ? singleAlphaInverseEffect(effects) : undefined
    const alphaInverseNodes = multiTreeAlphaInverseNodes || (alphaInverseNode ? [alphaInverseNode] : undefined)
    let alphaInverse: PptxAlphaInverse | undefined
    if (alphaInverseNodes) {
        let color: [number, number, number] | undefined
        for (const node of alphaInverseNodes) {
            const colorNode = children(node)[0]
            if (!colorNode) continue
            const channels = animationColorChannels(colorNode, theme, 'rgb')
            if (!channels) {
                warnings.add('An invalid picture alpha inverse color in an effect DAG was ignored; alpha-only inversion was applied.')
                continue
            }
            color = [channels[0] / 255, channels[1] / 255, channels[2] / 255]
        }
        const invertAlpha = alphaInverseNodes.length % 2 === 1
        if (invertAlpha || color) alphaInverse = {invertAlpha, color}
    }
    for (const effect of entries) {
        if (multiTreeAlphaInverseNodes || (effect === alphaInverseEntry && alphaInverseNode)) {
            continue
        } else if (effect.localName === 'outerShdw') {
            shadows.push(parsedShadow(effect, theme, warnings))
        } else if (effect.localName === 'innerShdw') {
            const shadow = parsedInnerShadow(effect, theme, warnings, target, supportsInset, supportsSvg)
            if (shadow) shadows.push(shadow)
        } else if (effect.localName === 'glow') {
            const glow = parsedGlow(effect, theme, warnings)
            if (glow) glows.push(glow)
        } else {
            warnings.add(effect.localName === 'cont'
                ? `Nested ${target} effect DAG containers are not rendered.`
                : `The ${target} ${effect.localName} effect in an effect DAG is not rendered.`)
        }
    }
    if (shadows.length > 1) warnings.add(target === 'text'
        ? 'Multiple text shadows use layered CSS text-shadow and SVG filters and have not been calibrated against Office/WPS.'
        : supportsSvg
            ? 'Multiple custom geometry and SVG polygon shadows use layered CSS drop-shadows and SVG filters and have not been calibrated against Office/WPS.'
            : `Multiple ${target} shadows use layered CSS box-shadows and have not been calibrated against Office/WPS.`)
    if (target !== 'text' && glows.length > 1) warnings.add(`Multiple ${target} glows use layered CSS drop-shadows and have not been calibrated against Office/WPS.`)
    return {shadow: shadows[0], shadows: shadows.length > 1 ? shadows : undefined, glow: glows[0], glows: glows.length > 1 ? glows : undefined, alphaInverse}
}

function referencedShapeEffects(shape: Element, theme: Record<string, string>, warnings: Set<string>, styles?: ParsedTheme, target: 'shape' | 'picture' = 'shape'): {shadow?: PptxShadow; shadows?: PptxShadow[]; glow?: PptxGlow; glows?: PptxGlow[]; alphaInverse?: PptxAlphaInverse} {
    const properties = child(shape, 'spPr')
    const geometry = child(properties, 'prstGeom')?.getAttribute('prst') || (child(properties, 'custGeom') ? 'custom' : 'rect')
    const supportsInset = target === 'picture' || (CSS_INSET_SHADOW_GEOMETRIES.has(geometry) && !isSvgPolygonGeometry(geometry))
    const supportsSvg = target === 'shape' && (geometry === 'custom' || isSvgPolygonGeometry(geometry))
    let effects = child(properties, 'effectLst') || child(properties, 'effectDag')
    if (!effects) {
        const reference = child(child(shape, 'style'), 'effectRef')
        if (!reference) return {}
        const raw = reference.getAttribute('idx') || ''
        const index = /^\d+$/.test(raw) ? Number(raw) : Number.NaN
        if (index === 0) return {}
        const style = Number.isSafeInteger(index) && index > 0 ? styles?.effectStyles[index - 1] : undefined
        if (!style) {
            warnings.add(`An invalid or missing ${target} effectRef style was rendered without effects.`)
            return {}
        }
        const color = readColor(children(reference)[0], theme)
        if (color) theme = {...theme, phClr: color}
        effects = child(style, 'effectLst') || child(style, 'effectDag')
        if (child(style, 'scene3d') || child(style, 'sp3d')) warnings.add(`Theme 3D ${target} effects are not rendered.`)
    }
    if (effects?.localName === 'effectDag') {
        return drawingEffectDag(effects, theme, warnings, target, supportsInset, supportsSvg)
    }
    const glows = glowEffects(effects, theme, warnings)
    const shadows = effectShadows(effects, theme, warnings, target, supportsInset, supportsSvg)
    for (const effect of children(effects)) if (!['outerShdw', 'innerShdw', 'glow'].includes(effect.localName)) warnings.add(`The ${target} ${effect.localName} effect is not rendered.`)
    if (shadows.length > 1) warnings.add(supportsSvg
        ? 'Multiple custom geometry and SVG polygon shadows use layered CSS drop-shadows and SVG filters and have not been calibrated against Office/WPS.'
        : `Multiple ${target} shadows use layered CSS box-shadows and have not been calibrated against Office/WPS.`)
    if (glows.length > 1) warnings.add(`Multiple ${target} glows use layered CSS drop-shadows and have not been calibrated against Office/WPS.`)
    return {shadow: shadows[0], shadows: shadows.length > 1 ? shadows : undefined, glow: glows[0], glows: glows.length > 1 ? glows : undefined}
}

function connectorGeometryPath(geometry: string, adjustments: Record<string, number>): string | undefined {
    const adjustment = (name: string) => Number(((adjustments[name] ?? 50_000) / 100).toFixed(3))
    if (geometry === 'bentConnector2') return 'M 0 0 L 1000 0 L 1000 1000'
    if (geometry === 'bentConnector3') {
        const x1 = adjustment('adj1')
        return `M 0 0 L ${x1} 0 L ${x1} 1000 L 1000 1000`
    }
    if (geometry === 'bentConnector4') {
        const x1 = adjustment('adj1')
        const y2 = adjustment('adj2')
        return `M 0 0 L ${x1} 0 L ${x1} ${y2} L 1000 ${y2} L 1000 1000`
    }
    if (geometry === 'bentConnector5') {
        const x1 = adjustment('adj1')
        const y2 = adjustment('adj2')
        const x3 = adjustment('adj3')
        return `M 0 0 L ${x1} 0 L ${x1} ${y2} L ${x3} ${y2} L ${x3} 1000 L 1000 1000`
    }
    if (geometry === 'curvedConnector2') return 'M 0 0 C 500 0 1000 500 1000 1000'
    if (geometry === 'curvedConnector3') {
        const midX = adjustment('adj1')
        return `M 0 0 C ${midX / 2} 0 ${midX} 250 ${midX} 500 C ${midX} 750 ${(1000 + midX) / 2} 1000 1000 1000`
    }
    if (geometry === 'curvedConnector4') {
        const midX = adjustment('adj1')
        const y4 = adjustment('adj2')
        const x1 = midX / 2
        const x3 = (1000 + midX) / 2
        const x4 = (midX + x3) / 2
        const x5 = (x3 + 1000) / 2
        const y1 = y4 / 2
        const y2 = y1 / 2
        const y3 = (y1 + y4) / 2
        const y5 = (1000 + y4) / 2
        return `M 0 0 C ${x1} 0 ${midX} ${y2} ${midX} ${y1} C ${midX} ${y3} ${x4} ${y4} ${x3} ${y4} C ${x5} ${y4} 1000 ${y5} 1000 1000`
    }
    if (geometry === 'curvedConnector5') {
        const x3 = adjustment('adj1')
        const y4 = adjustment('adj2')
        const x6 = adjustment('adj3')
        const x1 = (x3 + x6) / 2
        const x2 = x3 / 2
        const x4 = (x3 + x1) / 2
        const x5 = (x6 + x1) / 2
        const x7 = (x6 + 1000) / 2
        const y1 = y4 / 2
        const y2 = y1 / 2
        const y3 = (y1 + y4) / 2
        const y5 = (1000 + y4) / 2
        const y6 = (y5 + y4) / 2
        const y7 = (y5 + 1000) / 2
        return `M 0 0 C ${x2} 0 ${x3} ${y2} ${x3} ${y1} C ${x3} ${y3} ${x4} ${y4} ${x1} ${y4} C ${x5} ${y4} ${x6} ${y6} ${x6} ${y5} C ${x6} ${y7} ${x7} 1000 1000 1000`
    }
    return undefined
}

function parseShape(shape: Element, theme: Record<string, string>, warnings: Set<string>, inheritedStyle: PptxTextStyle = {}, inheritedParagraphStyles: ParagraphDefaultsByLevel = {}, warnPlaceholder = true, rels?: Map<string, PptxRelationship>, fillStyles?: ParsedTheme, groupFill?: string): PptxElement {
    const properties = child(shape, 'spPr')
    const customGeometry = child(properties, 'custGeom')
    const transform = shapeTransform(properties)
    const geometry = child(properties, 'prstGeom')?.getAttribute('prst') || (customGeometry ? 'custom' : 'rect')
    const isLineGeometry = LINE_GEOMETRIES.has(geometry)
    const presetGeometry = child(properties, 'prstGeom')
    const presetDefinition = geometry === 'cloud' ? getCloudPresetGeometry()
        : PRESET_GEOMETRIES_WITH_HANDWRITTEN_RENDERERS.has(geometry) ? undefined
            : getPresetShapeDefinition(geometry, presetGeometry, transform.width, transform.height)
    let presetAdjustment: number | undefined
    const shortSide = Math.min(transform.width, transform.height)
    const presetAdjustmentDefaults: Record<string, [number, number]> = {
        roundRect: [16667, 50000], round1Rect: [16667, 50000], triangle: [50000, 100000], parallelogram: [25000, 100000], trapezoid: [25000, 50000], hexagon: [25000, 50000], octagon: [29289, 50000], pentagon: [19098, 50000], star4: [12500, 50000], star5: [19098, 50000], star6: [28868, 50000], star7: [34601, 50000], star8: [37500, 50000], star10: [42533, 50000], star12: [37500, 50000], star16: [37500, 50000], star24: [37500, 50000], star32: [37500, 50000], plus: [25000, 50000], chevron: [50000, 100000], cube: [25000, 100000], plaque: [16667, 50000], snip1Rect: [16667, 50000], homePlate: [50000, 50000],
    }
    const presetGuideResolver = customGeometryGuideResolver(presetGeometry, transform.width, transform.height, warnings)
    presetAdjustmentDefaults.donut = [25000, 50000]
    const presetAdjustments = Object.fromEntries(children(child(presetGeometry, 'avLst'), 'gd').flatMap(node => {
        const name = node.getAttribute('name')
        const value = name ? presetGuideResolver(name) : undefined
        if (name && value === undefined && name !== 'adj') warnings.add(`The ${geometry} adjustment ${name} could not be evaluated; the preset default was used.`)
        return name && value !== undefined ? [[name, value] as const] : []
    }))
    if (geometry === 'bracketPair') {
        const adjustment = presetAdjustments.adj ?? 16_667
        presetAdjustments.adj = Math.max(0, Math.min(50_000, adjustment))
        if (presetAdjustments.adj !== adjustment) warnings.add('The bracketPair adjustment was constrained to its valid range.')
    }
    if (geometry === 'bracePair') {
        const adjustment = presetAdjustments.adj ?? 8_333
        presetAdjustments.adj = Math.max(0, Math.min(25_000, adjustment))
        if (presetAdjustments.adj !== adjustment) warnings.add('The bracePair adjustment was constrained to its valid range.')
    }
    if (geometry === 'leftBrace' || geometry === 'rightBrace') {
        const adjustment2 = presetAdjustments.adj2 ?? 50_000
        presetAdjustments.adj2 = Math.max(0, Math.min(100_000, adjustment2))
        if (presetAdjustments.adj2 !== adjustment2) warnings.add(`The ${geometry} adjustment adj2 was constrained to its valid range.`)
        const maximumAdjustment1 = shortSide > 0 ? Math.min(presetAdjustments.adj2, 100_000 - presetAdjustments.adj2) / 2 * transform.height / shortSide : 0
        const adjustment1 = presetAdjustments.adj1 ?? 8_333
        presetAdjustments.adj1 = Math.max(0, Math.min(maximumAdjustment1, adjustment1))
        if (presetAdjustments.adj1 !== adjustment1) warnings.add(`The ${geometry} adjustment adj1 was constrained to its valid range.`)
    }
    if (geometry === 'leftBracket' || geometry === 'rightBracket') {
        const maximumAdjustment = shortSide > 0 ? 50_000 * transform.height / shortSide : 50_000
        const adjustment = presetAdjustments.adj ?? 8_333
        presetAdjustments.adj = Math.max(0, Math.min(maximumAdjustment, adjustment))
        if (presetAdjustments.adj !== adjustment) warnings.add(`The ${geometry} adjustment was constrained to its valid range.`)
    }
    if (geometry === 'pie' || geometry === 'chord' || geometry === 'arc' || geometry === 'blockArc') {
        const defaults = geometry === 'pie' ? {adj1: 0, adj2: 16_200_000}
            : geometry === 'chord' ? {adj1: 2_700_000, adj2: 16_200_000}
                : geometry === 'arc' ? {adj1: 16_200_000, adj2: 0}
                    : {adj1: 10_800_000, adj2: 0}
        for (const name of ['adj1', 'adj2'] as const) {
            const value = presetAdjustments[name] ?? defaults[name]
            presetAdjustments[name] = Math.max(0, Math.min(21_599_999, value))
            if (presetAdjustments[name] !== value) warnings.add(`The ${geometry} adjustment ${name} was constrained to its valid angle range.`)
        }
    }
    if (geometry === 'blockArc') {
        const value = presetAdjustments.adj3 ?? 25_000
        presetAdjustments.adj3 = Math.max(0, Math.min(50_000, value))
        if (presetAdjustments.adj3 !== value) warnings.add('The blockArc adjustment adj3 was constrained to its valid thickness range.')
    }
    const adjustmentConfig = presetAdjustmentDefaults[geometry]
    if (adjustmentConfig) {
        const [defaultAdjustment, defaultMaximum] = adjustmentConfig
        const maximumAdjustment = defaultMaximum * (['parallelogram', 'trapezoid', 'hexagon', 'chevron', 'homePlate'].includes(geometry) ? transform.width / Math.min(transform.width, transform.height) : 1)
        const hasAdjustment = children(child(presetGeometry, 'avLst'), 'gd').some(node => node.getAttribute('name') === 'adj')
        const adjustment = hasAdjustment ? presetGuideResolver('adj') : defaultAdjustment
        if (adjustment === undefined) warnings.add(`An invalid ${geometry} adjustment uses the default geometry.`)
        presetAdjustment = Math.max(0, Math.min(maximumAdjustment, adjustment ?? defaultAdjustment)) / 100000
    }
    if (geometry === 'snip2DiagRect') {
        presetAdjustments.adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 0))
        presetAdjustments.adj2 = Math.max(0, Math.min(50000, presetAdjustments.adj2 ?? 16667))
    }
    if (geometry === 'snip2SameRect') {
        presetAdjustments.adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 16667))
        presetAdjustments.adj2 = Math.max(0, Math.min(50000, presetAdjustments.adj2 ?? 0))
    }
    if (geometry === 'round2DiagRect' || geometry === 'round2SameRect') {
        presetAdjustments.adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 16667))
        presetAdjustments.adj2 = Math.max(0, Math.min(50000, presetAdjustments.adj2 ?? 0))
    }
    if (geometry === 'snipRoundRect') {
        presetAdjustments.adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 16667))
        presetAdjustments.adj2 = Math.max(0, Math.min(50000, presetAdjustments.adj2 ?? 16667))
    }
    if (geometry === 'frame') presetAdjustments.adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 12500))
    if (geometry === 'halfFrame') {
        const maximumAdj2 = shortSide > 0 ? 100000 * transform.width / shortSide : 0
        const adj2 = Math.max(0, Math.min(maximumAdj2, presetAdjustments.adj2 ?? 33333))
        const x1 = shortSide * adj2 / 100000
        const maximumAdj1 = transform.width > 0 && shortSide > 0 ? 100000 * (transform.height - transform.height * x1 / transform.width) / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(maximumAdj1, presetAdjustments.adj1 ?? 33333))
        presetAdjustments.adj2 = adj2
    }
    if (geometry === 'nonIsoscelesTrapezoid') {
        const maximumAdjustment = shortSide > 0 ? 50000 * transform.width / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(maximumAdjustment, presetAdjustments.adj1 ?? 25000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumAdjustment, presetAdjustments.adj2 ?? 25000))
    }
    if (geometry === 'mathMultiply') presetAdjustments.adj1 = Math.max(0, Math.min(51965, presetAdjustments.adj1 ?? 23520))
    if (geometry === 'mathDivide') {
        const adj1 = Math.max(1000, Math.min(36745, presetAdjustments.adj1 ?? 23520))
        const maximumAdj3 = transform.height > 0 ? Math.min((73490 + adj1) / 4, 36745 * transform.width / transform.height) : 0
        const adj3 = Math.max(1000, Math.min(maximumAdj3, presetAdjustments.adj3 ?? 11760))
        presetAdjustments.adj1 = adj1
        presetAdjustments.adj3 = adj3
        presetAdjustments.adj2 = Math.max(0, Math.min(73490 - 4 * adj3 - adj1, presetAdjustments.adj2 ?? 5880))
    }
    if (geometry === 'mathNotEqual') {
        const adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 23520))
        presetAdjustments.adj1 = adj1
        presetAdjustments.adj2 = Math.max(4200000, Math.min(6600000, presetAdjustments.adj2 ?? 6600000))
        presetAdjustments.adj3 = Math.max(0, Math.min(100000 - 2 * adj1, presetAdjustments.adj3 ?? 11760))
    }
    if (geometry === 'leftRightArrow') {
        const maximumArrowWidth = shortSide > 0 ? 50000 * transform.width / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 50000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumArrowWidth, presetAdjustments.adj2 ?? 50000))
    }
    if (geometry === 'upDownArrow') {
        const maximumArrowHeight = shortSide > 0 ? 50000 * transform.height / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 50000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumArrowHeight, presetAdjustments.adj2 ?? 50000))
    }
    if (geometry === 'upArrow' || geometry === 'downArrow') {
        const maximumArrowDepth = shortSide > 0 ? 100000 * transform.height / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 50000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumArrowDepth, presetAdjustments.adj2 ?? 50000))
    }
    if (geometry === 'bentUpArrow') {
        presetAdjustments.adj1 = Math.max(0, Math.min(50000, presetAdjustments.adj1 ?? 25000))
        presetAdjustments.adj2 = Math.max(0, Math.min(50000, presetAdjustments.adj2 ?? 25000))
        presetAdjustments.adj3 = Math.max(0, Math.min(50000, presetAdjustments.adj3 ?? 25000))
    }
    if (geometry === 'quadArrow' || geometry === 'leftRightUpArrow') {
        const defaultAdjustment = geometry === 'quadArrow' ? 22500 : 25000
        presetAdjustments.adj2 = Math.max(0, Math.min(50000, presetAdjustments.adj2 ?? defaultAdjustment))
        const maximumAdj1 = 2 * presetAdjustments.adj2
        presetAdjustments.adj1 = Math.max(0, Math.min(maximumAdj1, presetAdjustments.adj1 ?? defaultAdjustment))
        const maximumAdj3 = (100000 - maximumAdj1) / 2
        presetAdjustments.adj3 = Math.max(0, Math.min(maximumAdj3, presetAdjustments.adj3 ?? defaultAdjustment))
    }
    if (geometry === 'rightArrow') {
        const maximumArrowLength = shortSide > 0 ? 100000 * transform.width / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 50000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumArrowLength, presetAdjustments.adj2 ?? 50000))
    }
    if (geometry === 'leftArrow') {
        const maximumArrowLength = shortSide > 0 ? 100000 * transform.width / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 50000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumArrowLength, presetAdjustments.adj2 ?? 50000))
    }
    if (geometry === 'notchedRightArrow') {
        const maximumArrowLength = shortSide > 0 ? 100000 * transform.width / shortSide : 0
        presetAdjustments.adj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 50000))
        presetAdjustments.adj2 = Math.max(0, Math.min(maximumArrowLength, presetAdjustments.adj2 ?? 50000))
    }
    const customPaths = customGeometry ? parseCustomGeometry(customGeometry, warnings) : presetDefinition ? parseCustomGeometry(presetDefinition, warnings) : undefined
    const linePath = isLineGeometry ? connectorGeometryPath(geometry, presetAdjustments) : undefined
    const chevronCut = shortSide * (presetAdjustment ?? 0.5)
    const chevronTextRect = transform.width > 2 * chevronCut
        ? {left: chevronCut / transform.width, top: 0, right: 1 - chevronCut / transform.width, bottom: 1} : undefined
    const hexagonMaximum = transform.width > 0 && shortSide > 0 ? 50000 * transform.width / shortSide : 0
    const hexagonBaseline = -hexagonMaximum / 2
    const hexagonOffset = (presetAdjustment ?? 0.25) * 100000
    const hexagonDelta = hexagonOffset + hexagonBaseline
    const hexagonScale = hexagonDelta > 0 ? 4 : 2
    const hexagonDenominator = hexagonDelta > 0 ? 3 : 2
    const hexagonStep = hexagonDelta > 0 ? hexagonBaseline : 0
    const hexagonRectInset = hexagonBaseline !== 0 ? (hexagonScale - (hexagonOffset + hexagonStep) / hexagonBaseline * hexagonDenominator / -1) / 24 : 0
    const hexagonTextRect = hexagonRectInset >= 0 && hexagonRectInset < 0.5
        ? {left: hexagonRectInset, top: hexagonRectInset, right: 1 - hexagonRectInset, bottom: 1 - hexagonRectInset} : undefined
    const trapezoidTextInset = transform.width > 0 && shortSide > 0 ? (presetAdjustment ?? 0.25) * shortSide / transform.width * 2 / 3 : 0
    const octagonTextInset = shortSide * (presetAdjustment ?? 0.29289) / 2
    const octagonTextRect = transform.width > 0 && transform.height > 0
        ? {left: octagonTextInset / transform.width, top: octagonTextInset / transform.height, right: 1 - octagonTextInset / transform.width, bottom: 1 - octagonTextInset / transform.height} : undefined
    const pentagonHorizontalFactor = presetAdjustments.hf ?? 105146
    const pentagonVerticalFactor = presetAdjustments.vf ?? 110557
    const pentagonHalfWidth = transform.width * pentagonHorizontalFactor / 200000
    const pentagonHalfHeight = transform.height * pentagonVerticalFactor / 200000
    const pentagonDx1 = pentagonHalfWidth * Math.cos(Math.PI / 10)
    const pentagonDx2 = pentagonHalfWidth * Math.cos(17 * Math.PI / 10)
    const pentagonDy1 = pentagonHalfHeight * Math.sin(Math.PI / 10)
    const pentagonDy2 = pentagonHalfHeight * Math.sin(17 * Math.PI / 10)
    const pentagonX2 = transform.width / 2 - pentagonDx2
    const pentagonX3 = transform.width / 2 + pentagonDx2
    const pentagonY1 = transform.height * pentagonVerticalFactor / 200000 - pentagonDy1
    const pentagonY2 = transform.height * pentagonVerticalFactor / 200000 - pentagonDy2
    const pentagonTextTop = pentagonDx1 !== 0 ? pentagonY1 * pentagonDx2 / pentagonDx1 : 0
    const pentagonTextRect = transform.width > 0 && transform.height > 0 && pentagonDx1 !== 0
        ? {left: pentagonX2 / transform.width, top: pentagonTextTop / transform.height, right: pentagonX3 / transform.width, bottom: pentagonY2 / transform.height} : undefined
    const star5Adjustment = (presetAdjustment ?? 0.19098) * 2
    const star5HalfWidth = transform.width / 2 * pentagonHorizontalFactor / 100000
    const star5HalfHeight = transform.height / 2 * pentagonVerticalFactor / 100000
    const star5CenterY = transform.height / 2 * pentagonVerticalFactor / 100000
    const star5InnerHalfWidth = star5HalfWidth * star5Adjustment
    const star5InnerHalfHeight = star5HalfHeight * star5Adjustment
    const star5Sdx1 = star5InnerHalfWidth * Math.cos(19 * Math.PI / 10)
    const star5Sdy1 = star5InnerHalfHeight * Math.sin(3 * Math.PI / 10)
    const star5TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star5Sdx1) / transform.width, top: (star5CenterY - star5Sdy1) / transform.height, right: (transform.width / 2 + star5Sdx1) / transform.width, bottom: (star5CenterY + star5InnerHalfHeight) / transform.height} : undefined
    const star6Adjustment = (presetAdjustment ?? 0.28868) * 2
    const star6HalfWidth = transform.width / 2 * (presetAdjustments.hf ?? 115470) / 100000
    const star6InnerHalfWidth = star6HalfWidth * star6Adjustment
    const star6InnerHalfHeight = transform.height / 2 * star6Adjustment
    const star6TextInsetY = star6InnerHalfHeight * Math.sin(Math.PI / 3)
    const star6TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star6InnerHalfWidth) / transform.width, top: (transform.height / 2 - star6TextInsetY) / transform.height, right: (transform.width / 2 + star6InnerHalfWidth) / transform.width, bottom: (transform.height / 2 + star6TextInsetY) / transform.height} : undefined
    const star8Adjustment = (presetAdjustment ?? 0.375) * 2
    const star8InnerHalfWidth = transform.width / 2 * star8Adjustment
    const star8InnerHalfHeight = transform.height / 2 * star8Adjustment
    const star8TextInsetX = star8InnerHalfWidth * 0.92388
    const star8TextInsetY = star8InnerHalfHeight * 0.92388
    const star8TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star8TextInsetX) / transform.width, top: (transform.height / 2 - star8TextInsetY) / transform.height, right: (transform.width / 2 + star8TextInsetX) / transform.width, bottom: (transform.height / 2 + star8TextInsetY) / transform.height} : undefined
    const star4Adjustment = (presetAdjustment ?? 0.125) * 2
    const star4HalfWidth = transform.width / 2 * star4Adjustment
    const star4HalfHeight = transform.height / 2 * star4Adjustment
    const star4TextInsetX = star4HalfWidth * Math.cos(Math.PI / 4)
    const star4TextInsetY = star4HalfHeight * Math.sin(Math.PI / 4)
    const star4TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star4TextInsetX) / transform.width, top: (transform.height / 2 - star4TextInsetY) / transform.height, right: (transform.width / 2 + star4TextInsetX) / transform.width, bottom: (transform.height / 2 + star4TextInsetY) / transform.height} : undefined
    const star7InnerScale = (presetAdjustment ?? 0.34601) * 2
    const star7HorizontalFactor = presetAdjustments.hf ?? 102572
    const star7VerticalFactor = presetAdjustments.vf ?? 105210
    const star7OuterHalfWidth = transform.width / 2 * star7HorizontalFactor / 100000
    const star7InnerHalfWidth = star7OuterHalfWidth * star7InnerScale
    const star7InnerHalfHeight = transform.height / 2 * star7VerticalFactor / 100000 * star7InnerScale
    const star7CenterY = transform.height / 2 * star7VerticalFactor / 100000
    const star7TextInsetX = star7InnerHalfWidth * 0.78183
    const star7TextInsetTop = star7InnerHalfHeight * 0.90097
    const star7TextInsetBottom = star7InnerHalfHeight * 0.62349
    const star7TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star7TextInsetX) / transform.width, top: (star7CenterY - star7TextInsetTop) / transform.height, right: (transform.width / 2 + star7TextInsetX) / transform.width, bottom: (star7CenterY + star7TextInsetBottom) / transform.height} : undefined
    // ECMA-376 Part 1 defines the star10/star12 preset points and text rectangles with these guides.
    const star10InnerScale = (presetAdjustment ?? 0.42533) * 2
    const star10OuterHalfWidth = transform.width / 2 * (presetAdjustments.hf ?? 105146) / 100000
    const star10InnerHalfWidth = star10OuterHalfWidth * star10InnerScale
    const star10InnerHalfHeight = transform.height / 2 * star10InnerScale
    const star10TextInsetX = star10InnerHalfWidth * 0.80902
    const star10TextInsetY = star10InnerHalfHeight * 0.58779
    const star10TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star10TextInsetX) / transform.width, top: (transform.height / 2 - star10TextInsetY) / transform.height, right: (transform.width / 2 + star10TextInsetX) / transform.width, bottom: (transform.height / 2 + star10TextInsetY) / transform.height} : undefined
    const star12InnerScale = (presetAdjustment ?? 0.375) * 2
    const star12InnerHalfWidth = transform.width / 2 * star12InnerScale
    const star12InnerHalfHeight = transform.height / 2 * star12InnerScale
    const star12TextInsetX = star12InnerHalfWidth * Math.cos(Math.PI / 4)
    const star12TextInsetY = star12InnerHalfHeight * Math.cos(Math.PI / 4)
    const star12TextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - star12TextInsetX) / transform.width, top: (transform.height / 2 - star12TextInsetY) / transform.height, right: (transform.width / 2 + star12TextInsetX) / transform.width, bottom: (transform.height / 2 + star12TextInsetY) / transform.height} : undefined
    // ECMA-376 Part 1 star16/star24/star32 definitions share a 45-degree inset text rectangle.
    const highPointStarInnerScale = (presetAdjustment ?? 0.375) * 2
    const highPointStarInnerHalfWidth = transform.width / 2 * highPointStarInnerScale
    const highPointStarInnerHalfHeight = transform.height / 2 * highPointStarInnerScale
    const highPointStarTextInsetX = highPointStarInnerHalfWidth * Math.SQRT1_2
    const highPointStarTextInsetY = highPointStarInnerHalfHeight * Math.SQRT1_2
    const highPointStarTextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - highPointStarTextInsetX) / transform.width, top: (transform.height / 2 - highPointStarTextInsetY) / transform.height, right: (transform.width / 2 + highPointStarTextInsetX) / transform.width, bottom: (transform.height / 2 + highPointStarTextInsetY) / transform.height} : undefined
    const plusInset = shortSide * (presetAdjustment ?? 0.25)
    const plusTextRect = transform.width > 0 && transform.height > 0
        ? transform.width > transform.height
            ? {left: 0, top: plusInset / transform.height, right: 1, bottom: 1 - plusInset / transform.height}
            : {left: plusInset / transform.width, top: 0, right: 1 - plusInset / transform.width, bottom: 1}
        : undefined
    const mathHalfBarWidth = transform.width * 73490 / 200000
    const mathMinusAdj1 = Math.max(0, Math.min(100000, presetAdjustments.adj1 ?? 23520))
    const mathMinusHalfBarHeight = transform.height * mathMinusAdj1 / 200000
    const mathMinusTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0.5 - mathHalfBarWidth / transform.width, top: 0.5 - mathMinusHalfBarHeight / transform.height, right: 0.5 + mathHalfBarWidth / transform.width, bottom: 0.5 + mathMinusHalfBarHeight / transform.height}
        : undefined
    const mathEqualAdj1 = Math.max(0, Math.min(36745, presetAdjustments.adj1 ?? 23520))
    const mathEqualAdj2 = Math.max(0, Math.min(100000 - 2 * mathEqualAdj1, presetAdjustments.adj2 ?? 11760))
    const mathEqualTextInsetY = mathEqualAdj1 / 100000 + mathEqualAdj2 / 200000
    const mathEqualTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0.5 - mathHalfBarWidth / transform.width, top: 0.5 - mathEqualTextInsetY, right: 0.5 + mathHalfBarWidth / transform.width, bottom: 0.5 + mathEqualTextInsetY}
        : undefined
    const mathPlusAdj1 = Math.max(0, Math.min(73490, presetAdjustments.adj1 ?? 23520))
    const mathPlusCrossHalfArm = shortSide * mathPlusAdj1 / 200000
    const mathPlusTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0.5 - mathHalfBarWidth / transform.width, top: 0.5 - mathPlusCrossHalfArm / transform.height, right: 0.5 + mathHalfBarWidth / transform.width, bottom: 0.5 + mathPlusCrossHalfArm / transform.height}
        : undefined
    const mathDivideAdj1 = presetAdjustments.adj1 ?? 23520
    const mathDivideTextHalfHeight = transform.height * mathDivideAdj1 / 200000
    const mathDivideTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0.5 - mathHalfBarWidth / transform.width, top: 0.5 - mathDivideTextHalfHeight / transform.height, right: 0.5 + mathHalfBarWidth / transform.width, bottom: 0.5 + mathDivideTextHalfHeight / transform.height}
        : undefined
    const mathMultiplyDimensionsValid = transform.width > 0 && transform.height > 0
    const mathMultiplyAngle = mathMultiplyDimensionsValid ? Math.atan2(transform.height, transform.width) : 0
    const mathMultiplyLength = mathMultiplyDimensionsValid ? Math.hypot(transform.width, transform.height) * (1 - 0.51965) : 0
    const mathMultiplyX = Math.cos(mathMultiplyAngle) * mathMultiplyLength / 2
    const mathMultiplyY = Math.sin(mathMultiplyAngle) * mathMultiplyLength / 2
    const mathMultiplyThickness = shortSide * (presetAdjustments.adj1 ?? 23520) / 100000
    const mathMultiplyDx = Math.sin(mathMultiplyAngle) * mathMultiplyThickness / 2
    const mathMultiplyDy = Math.cos(mathMultiplyAngle) * mathMultiplyThickness / 2
    const mathMultiplyXA = mathMultiplyX - mathMultiplyDx
    const mathMultiplyYB = mathMultiplyY - mathMultiplyDy
    const mathMultiplyTextRect = mathMultiplyDimensionsValid
        ? {left: mathMultiplyXA / transform.width, top: mathMultiplyYB / transform.height, right: (transform.width - mathMultiplyXA) / transform.width, bottom: (transform.height - mathMultiplyYB) / transform.height}
        : undefined
    const mathNotEqualAdj1 = presetAdjustments.adj1 ?? 23520
    const mathNotEqualAdj3 = presetAdjustments.adj3 ?? 11760
    const mathNotEqualTextInset = transform.height * (mathNotEqualAdj1 / 100000 + mathNotEqualAdj3 / 200000)
    const mathNotEqualTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0.13255, top: 0.5 - mathNotEqualTextInset / transform.height, right: 0.86745, bottom: 0.5 + mathNotEqualTextInset / transform.height}
        : undefined
    const snip1Inset = shortSide * (presetAdjustment ?? 0.16667)
    const snip1RectTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0, top: snip1Inset / 2 / transform.height, right: 1 - snip1Inset / 2 / transform.width, bottom: 1}
        : undefined
    const snip2DiagInset = Math.max(shortSide * (presetAdjustments.adj1 ?? 0) / 100000, shortSide * (presetAdjustments.adj2 ?? 16667) / 100000) / 2
    const snip2DiagRectTextRect = transform.width > 0 && transform.height > 0
        ? {left: snip2DiagInset / transform.width, top: snip2DiagInset / transform.height, right: 1 - snip2DiagInset / transform.width, bottom: 1 - snip2DiagInset / transform.height}
        : undefined
    const snip2SameTopInset = shortSide * (presetAdjustments.adj1 ?? 16667) / 100000
    const snip2SameBottomInset = shortSide * (presetAdjustments.adj2 ?? 0) / 100000
    const snip2SameTextSideInset = Math.max(snip2SameTopInset, snip2SameBottomInset) / 2
    const snip2SameRectTextRect = transform.width > 0 && transform.height > 0
        ? {left: snip2SameTextSideInset / transform.width, top: snip2SameTopInset / 2 / transform.height, right: 1 - snip2SameTextSideInset / transform.width, bottom: 1 - snip2SameBottomInset / 2 / transform.height}
        : undefined
    const arrowX2 = shortSide * (presetAdjustments.adj2 ?? 50000) / 100000
    const arrowY1 = transform.height / 2 - transform.height * (presetAdjustments.adj1 ?? 50000) / 200000
    const arrowDx1 = transform.height > 0 ? arrowY1 * arrowX2 / (transform.height / 2) : 0
    const leftRightArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: (arrowX2 - arrowDx1) / transform.width, top: arrowY1 / transform.height, right: (transform.width - arrowX2 + arrowDx1) / transform.width, bottom: (transform.height - arrowY1) / transform.height}
        : undefined
    const upDownArrowX1 = transform.width / 2 - transform.width * (presetAdjustments.adj1 ?? 50000) / 200000
    const upDownArrowY2 = shortSide * (presetAdjustments.adj2 ?? 50000) / 100000
    const upDownArrowDy1 = transform.width > 0 ? upDownArrowX1 * upDownArrowY2 / (transform.width / 2) : 0
    const upDownArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: upDownArrowX1 / transform.width, top: (upDownArrowY2 - upDownArrowDy1) / transform.height, right: (transform.width - upDownArrowX1) / transform.width, bottom: (transform.height - upDownArrowY2 + upDownArrowDy1) / transform.height}
        : undefined
    const verticalArrowX1 = transform.width / 2 - transform.width * (presetAdjustments.adj1 ?? 50000) / 200000
    const verticalArrowX2 = transform.width - verticalArrowX1
    const verticalArrowDepth = shortSide * (presetAdjustments.adj2 ?? 50000) / 100000
    const upArrowY2 = verticalArrowDepth
    const upArrowDy1 = transform.width > 0 ? verticalArrowX1 * upArrowY2 / (transform.width / 2) : 0
    const upArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: verticalArrowX1 / transform.width, top: (upArrowY2 - upArrowDy1) / transform.height, right: verticalArrowX2 / transform.width, bottom: 1}
        : undefined
    const downArrowY1 = transform.height - verticalArrowDepth
    const downArrowDy2 = transform.width > 0 ? verticalArrowX1 * verticalArrowDepth / (transform.width / 2) : 0
    const downArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: verticalArrowX1 / transform.width, top: 0, right: verticalArrowX2 / transform.width, bottom: (downArrowY1 + downArrowDy2) / transform.height}
        : undefined
    const bentUpArrowA1 = presetAdjustments.adj1 ?? 25000
    const bentUpArrowA2 = presetAdjustments.adj2 ?? 25000
    const bentUpArrowA3 = presetAdjustments.adj3 ?? 25000
    const bentUpArrowDx2 = shortSide * bentUpArrowA1 / 200000
    const bentUpArrowX3 = transform.width - shortSide * bentUpArrowA1 / 100000
    const bentUpArrowX4 = bentUpArrowX3 + bentUpArrowDx2
    const bentUpArrowY2 = transform.height - shortSide * bentUpArrowA1 / 100000
    const bentUpArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0, top: bentUpArrowY2 / transform.height, right: bentUpArrowX4 / transform.width, bottom: 1}
        : undefined
    const quadArrowA1 = presetAdjustments.adj1 ?? 22500
    const quadArrowA2 = presetAdjustments.adj2 ?? 22500
    const quadArrowA3 = presetAdjustments.adj3 ?? 22500
    const quadArrowX1 = shortSide * quadArrowA3 / 100000
    const quadArrowDx2 = shortSide * quadArrowA2 / 100000
    const quadArrowDx3 = shortSide * quadArrowA1 / 200000
    const quadArrowInset = quadArrowDx2 > 0 ? quadArrowDx3 * quadArrowX1 / quadArrowDx2 : 0
    const quadArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: quadArrowInset / transform.width, top: (transform.height / 2 - quadArrowDx3) / transform.height, right: (transform.width - quadArrowInset) / transform.width, bottom: (transform.height / 2 + quadArrowDx3) / transform.height}
        : undefined
    const leftRightUpArrowA1 = presetAdjustments.adj1 ?? 25000
    const leftRightUpArrowA2 = presetAdjustments.adj2 ?? 25000
    const leftRightUpArrowA3 = presetAdjustments.adj3 ?? 25000
    const leftRightUpArrowX1 = shortSide * leftRightUpArrowA3 / 100000
    const leftRightUpArrowDx2 = shortSide * leftRightUpArrowA2 / 100000
    const leftRightUpArrowDx3 = shortSide * leftRightUpArrowA1 / 200000
    const leftRightUpArrowInset = leftRightUpArrowDx2 > 0 ? leftRightUpArrowDx3 * leftRightUpArrowX1 / leftRightUpArrowDx2 : 0
    const leftRightUpArrowY3 = transform.height - leftRightUpArrowDx2 - leftRightUpArrowDx3
    const leftRightUpArrowY5 = transform.height - leftRightUpArrowDx2 + leftRightUpArrowDx3
    const leftRightUpArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: leftRightUpArrowInset / transform.width, top: leftRightUpArrowY3 / transform.height, right: (transform.width - leftRightUpArrowInset) / transform.width, bottom: leftRightUpArrowY5 / transform.height}
        : undefined
    const rightArrowLength = shortSide * (presetAdjustments.adj2 ?? 50000) / 100000
    const rightArrowX1 = transform.width - rightArrowLength
    const rightArrowY1 = transform.height / 2 - transform.height * (presetAdjustments.adj1 ?? 50000) / 200000
    const rightArrowDx2 = transform.height > 0 ? rightArrowY1 * rightArrowLength / (transform.height / 2) : 0
    const rightArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0, top: rightArrowY1 / transform.height, right: (rightArrowX1 + rightArrowDx2) / transform.width, bottom: (transform.height - rightArrowY1) / transform.height}
        : undefined
    const leftArrowLength = shortSide * (presetAdjustments.adj2 ?? 50000) / 100000
    const leftArrowY1 = transform.height / 2 - transform.height * (presetAdjustments.adj1 ?? 50000) / 200000
    const leftArrowDx1 = transform.height > 0 ? leftArrowY1 * leftArrowLength / (transform.height / 2) : 0
    const leftArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: (leftArrowLength - leftArrowDx1) / transform.width, top: leftArrowY1 / transform.height, right: 1, bottom: (transform.height - leftArrowY1) / transform.height}
        : undefined
    const notchedRightArrowLength = shortSide * (presetAdjustments.adj2 ?? 50000) / 100000
    const notchedRightArrowX2 = transform.width - notchedRightArrowLength
    const notchedRightArrowHalfHeight = transform.height * (presetAdjustments.adj1 ?? 50000) / 200000
    const notchedRightArrowY1 = transform.height / 2 - notchedRightArrowHalfHeight
    const notchedRightArrowX1 = transform.height > 0 ? notchedRightArrowHalfHeight * notchedRightArrowLength / (transform.height / 2) : 0
    const notchedRightArrowTextRect = transform.width > 0 && transform.height > 0
        ? {left: notchedRightArrowX1 / transform.width, top: notchedRightArrowY1 / transform.height, right: (transform.width - notchedRightArrowX1) / transform.width, bottom: (transform.height - notchedRightArrowY1) / transform.height}
        : undefined
    const dodecagonTextInset = 2894 / 21600
    const magneticTapeTextInset = (1 - Math.SQRT1_2) / 2
    const homePlateCut = shortSide * (presetAdjustment ?? 0.5)
    const homePlateTextRect = transform.width > 0 ? {left: 0, top: 0, right: (transform.width - homePlateCut / 2) / transform.width, bottom: 1} : undefined
    const cubeInset = shortSide * (presetAdjustment ?? 0.25)
    const cubeTextRect = transform.width > 0 && transform.height > 0
        ? {left: 0, top: cubeInset / transform.height, right: (transform.width - cubeInset) / transform.width, bottom: 1}
        : undefined
    const plaqueRadius = shortSide * (presetAdjustment ?? 0.16667)
    const plaqueInset = plaqueRadius * 0.70711
    const plaqueTextRect = transform.width > 0 && transform.height > 0
        ? {left: plaqueInset / transform.width, top: plaqueInset / transform.height, right: 1 - plaqueInset / transform.width, bottom: 1 - plaqueInset / transform.height}
        : undefined
    const frameInset = shortSide * (presetAdjustments.adj1 ?? 12500) / 100000
    const frameTextRect = transform.width > 0 && transform.height > 0
        ? {left: frameInset / transform.width, top: frameInset / transform.height, right: 1 - frameInset / transform.width, bottom: 1 - frameInset / transform.height}
        : undefined
    const tabInset = Math.hypot(transform.width, transform.height) / 20
    const tabTextRect = transform.width > 0 && transform.height > 0
        ? {left: tabInset / transform.width, top: tabInset / transform.height, right: 1 - tabInset / transform.width, bottom: 1 - tabInset / transform.height}
        : undefined
    const round1RectRadius = shortSide * (presetAdjustment ?? 0.16667)
    const donutTextInset = (1 - Math.SQRT1_2) / 2
    const donutTextRect = {left: donutTextInset, top: donutTextInset, right: 1 - donutTextInset, bottom: 1 - donutTextInset}
    const pieWedgeTextInset = 1 - Math.SQRT1_2
    const pieWedgeTextRect = {left: pieWedgeTextInset, top: pieWedgeTextInset, right: 1, bottom: 1}
    const angleArcTextRect = (() => {
        if (!['arc', 'blockArc'].includes(geometry) || !(transform.width > 0) || !(transform.height > 0)) return undefined
        const fullTurn = 21_600_000
        const startAngle = presetAdjustments.adj1 ?? (geometry === 'blockArc' ? 10_800_000 : 16_200_000)
        const endAngle = presetAdjustments.adj2 ?? 0
        let sweepAngle = endAngle - startAngle
        if (sweepAngle <= 0) sweepAngle += fullTurn
        const [startX, startY] = presetEllipsePoint(transform.width, transform.height, startAngle)
        const [endX, endY] = presetEllipsePoint(transform.width, transform.height, endAngle)
        let left = Math.min(startX, endX) / 1000, right = Math.max(startX, endX) / 1000
        let top = Math.min(startY, endY) / 1000, bottom = Math.max(startY, endY) / 1000
        const includesAngle = (angle: number) => sweepAngle === fullTurn || sweepAngle > (angle - startAngle + fullTurn) % fullTurn
        if (includesAngle(0)) right = 1
        if (includesAngle(9_000_000)) bottom = 1
        if (includesAngle(10_800_000)) left = 0
        if (includesAngle(16_200_000)) top = 0
        return {left, top, right, bottom}
    })()
    const round1RectTextRect = transform.width > 0
        ? {left: 0, top: 0, right: 1 - round1RectRadius * 0.29289 / transform.width, bottom: 1}
        : undefined
    const round2DiagRectRadius1 = shortSide * (presetAdjustments.adj1 ?? 16667) / 100000
    const round2DiagRectRadius2 = shortSide * (presetAdjustments.adj2 ?? 0) / 100000
    const round2DiagRectInset = Math.max(round2DiagRectRadius1, round2DiagRectRadius2) * 0.29289
    const round2DiagRectTextRect = transform.width > 0 && transform.height > 0
        ? {left: round2DiagRectInset / transform.width, top: round2DiagRectInset / transform.height, right: 1 - round2DiagRectInset / transform.width, bottom: 1 - round2DiagRectInset / transform.height}
        : undefined
    const round2SameRectRadius1 = shortSide * (presetAdjustments.adj1 ?? 16667) / 100000
    const round2SameRectRadius2 = shortSide * (presetAdjustments.adj2 ?? 0) / 100000
    const round2SameRectInsetTop = round2SameRectRadius1 * 0.29289
    const round2SameRectInsetBottom = round2SameRectRadius2 * 0.29289
    const round2SameRectInsetSide = Math.max(round2SameRectInsetTop, round2SameRectInsetBottom)
    const round2SameRectTextRect = transform.width > 0 && transform.height > 0
        ? {left: round2SameRectInsetSide / transform.width, top: round2SameRectInsetTop / transform.height, right: 1 - round2SameRectInsetSide / transform.width, bottom: 1 - round2SameRectInsetBottom / transform.height}
        : undefined
    const snipRoundRectRadius = shortSide * (presetAdjustments.adj1 ?? 16667) / 100000
    const snipRoundRectCut = shortSide * (presetAdjustments.adj2 ?? 16667) / 100000
    const snipRoundRectInset = snipRoundRectRadius * 0.29289
    const snipRoundRectTextRect = transform.width > 0 && transform.height > 0
        ? {left: snipRoundRectInset / transform.width, top: snipRoundRectInset / transform.height, right: (transform.width - snipRoundRectCut / 2) / transform.width, bottom: 1}
        : undefined
    const nonIsoscelesTrapezoidMax = shortSide > 0 ? 50000 * transform.width / shortSide : 0
    const nonIsoscelesTrapezoidAdj1 = presetAdjustments.adj1 ?? 25000
    const nonIsoscelesTrapezoidAdj2 = presetAdjustments.adj2 ?? 25000
    const nonIsoscelesTrapezoidTextRect = transform.width > 0 && transform.height > 0 && nonIsoscelesTrapezoidMax > 0
        ? {left: (transform.width / 3 * nonIsoscelesTrapezoidAdj1 / nonIsoscelesTrapezoidMax) / transform.width, top: (transform.height / 3 * Math.max(nonIsoscelesTrapezoidAdj1, nonIsoscelesTrapezoidAdj2) / nonIsoscelesTrapezoidMax) / transform.height, right: (transform.width - transform.width / 3 * nonIsoscelesTrapezoidAdj2 / nonIsoscelesTrapezoidMax) / transform.width, bottom: 1}
        : undefined
    const heptagonHalfWidth = transform.width / 2 * (presetAdjustments.hf ?? 102572) / 100000
    const heptagonHalfHeight = transform.height / 2 * (presetAdjustments.vf ?? 105210) / 100000
    const heptagonY1 = heptagonHalfHeight - heptagonHalfHeight * 62349 / 100000
    const heptagonX2 = transform.width / 2 - heptagonHalfWidth * 78183 / 100000
    const heptagonTextRect = transform.width > 0 && transform.height > 0
        ? {left: heptagonX2 / transform.width, top: heptagonY1 / transform.height, right: (transform.width - heptagonX2) / transform.width, bottom: (transform.height - heptagonY1) / transform.height}
        : undefined
    const decagonHalfHeight = transform.height / 2 * (presetAdjustments.vf ?? 105146) / 100000
    const decagonXInset = transform.width / 2 * Math.cos(Math.PI / 5)
    const decagonYInset = decagonHalfHeight * Math.sin(Math.PI / 5)
    const decagonTextRect = transform.width > 0 && transform.height > 0
        ? {left: (transform.width / 2 - decagonXInset) / transform.width, top: (transform.height / 2 - decagonYInset) / transform.height, right: (transform.width / 2 + decagonXInset) / transform.width, bottom: (transform.height / 2 + decagonYInset) / transform.height}
        : undefined
    const bracketTextRect = ['leftBracket', 'rightBracket'].includes(geometry) && transform.width > 0 && transform.height > 0 && shortSide > 0
        ? (() => {
            const y1 = shortSide * (presetAdjustments.adj ?? 8_333) / 100_000 / transform.height
            const inset = Math.SQRT1_2
            return {
                left: geometry === 'leftBracket' ? 1 - inset : 0,
                top: y1 * (1 - inset),
                right: geometry === 'leftBracket' ? 1 : inset,
                bottom: 1 - y1 * (1 + inset)
            }
        })()
        : undefined
    const bracketPairTextRect = geometry === 'bracketPair' && transform.width > 0 && transform.height > 0 && shortSide > 0
        ? (() => {
            const inset = shortSide * (presetAdjustments.adj ?? 16_667) / 100_000 * 29_289 / 100_000
            return {left: inset / transform.width, top: inset / transform.height, right: 1 - inset / transform.width, bottom: 1 - inset / transform.height}
        })()
        : undefined
    const bracePairTextRect = geometry === 'bracePair' && transform.width > 0 && transform.height > 0 && shortSide > 0
        ? (() => {
            const radius = shortSide * (presetAdjustments.adj ?? 8_333) / 100_000
            const topInset = radius * 29_289 / 100_000
            const sideInset = radius + topInset
            return {left: sideInset / transform.width, top: sideInset / transform.height, right: 1 - sideInset / transform.width, bottom: 1 - topInset / transform.height}
        })()
        : undefined
    const singleBraceTextRect = (geometry === 'leftBrace' || geometry === 'rightBrace') && transform.width > 0 && transform.height > 0 && shortSide > 0
        ? (() => {
            const radius = shortSide * (presetAdjustments.adj1 ?? 8_333) / 100_000
            const horizontalInset = transform.width / 2 * Math.SQRT1_2
            const verticalInset = radius * (1 - Math.SQRT1_2)
            return {
                left: geometry === 'leftBrace' ? 1 - horizontalInset / transform.width : 0,
                top: verticalInset / transform.height,
                right: geometry === 'leftBrace' ? 1 : horizontalInset / transform.width,
                bottom: 1 - verticalInset / transform.height
            }
        })()
        : undefined
    // ECMA-376 presetShapeDefinitions.xml defines ellipse text bounds with center ± cos/sin(45°).
    const customTextRect = customGeometry ? parseCustomGeometryTextRect(customGeometry, warnings)
        : presetDefinition ? parseCustomGeometryTextRect(presetDefinition, warnings)
        : geometry === 'ellipse' ? {left: (1 - Math.SQRT1_2) / 2, top: (1 - Math.SQRT1_2) / 2, right: (1 + Math.SQRT1_2) / 2, bottom: (1 + Math.SQRT1_2) / 2}
        : geometry === 'heart' ? {left: 1 / 6, top: 1 / 4, right: 5 / 6, bottom: 2 / 3}
        : geometry === 'lightningBolt' ? {left: 8757 / 21600, top: 7437 / 21600, right: 13917 / 21600, bottom: 14277 / 21600}
        : ['leftBracket', 'rightBracket'].includes(geometry) ? bracketTextRect
        : geometry === 'bracketPair' ? bracketPairTextRect
        : geometry === 'bracePair' ? bracePairTextRect
        : geometry === 'leftBrace' || geometry === 'rightBrace' ? singleBraceTextRect
        : geometry === 'triangle' ? {left: (presetAdjustment ?? 0.5) / 2, top: 0.5, right: (presetAdjustment ?? 0.5) / 2 + 0.5, bottom: 1}
        : ['diamond', 'flowChartDecision', 'flowChartSort', 'flowChartCollate'].includes(geometry) ? {left: 0.25, top: 0.25, right: 0.75, bottom: 0.75}
        : geometry === 'flowChartOffpageConnector' ? {left: 0, top: 0, right: 1, bottom: 0.8}
        : geometry === 'flowChartDocument' ? {left: 0, top: 0, right: 1, bottom: 17322 / 21600}
        : ['flowChartManualInput', 'flowChartPunchedCard'].includes(geometry) ? {left: 0, top: 0.2, right: 1, bottom: 1}
        : geometry === 'flowChartManualOperation' ? {left: 0.2, top: 0, right: 0.8, bottom: 1}
        : geometry === 'flowChartInputOutput' ? {left: 0.2, top: 0, right: 0.8, bottom: 1}
        : geometry === 'flowChartPreparation' ? {left: 0.2, top: 0, right: 0.8, bottom: 1}
        : geometry === 'flowChartTerminator' ? {left: 1018 / 21600, top: 3163 / 21600, right: 20582 / 21600, bottom: 18437 / 21600}
        : geometry === 'flowChartPredefinedProcess' ? {left: 1 / 8, top: 0, right: 7 / 8, bottom: 1}
        : geometry === 'flowChartExtract' ? {left: 0.25, top: 0.5, right: 0.75, bottom: 1}
        : geometry === 'flowChartMerge' ? {left: 0.25, top: 0, right: 0.75, bottom: 0.5}
        : ['flowChartConnector', 'flowChartSummingJunction', 'flowChartOr'].includes(geometry) ? {left: (1 - Math.SQRT1_2) / 2, top: (1 - Math.SQRT1_2) / 2, right: (1 + Math.SQRT1_2) / 2, bottom: (1 + Math.SQRT1_2) / 2}
        : geometry === 'flowChartDelay' ? {left: 0, top: (1 - Math.SQRT1_2) / 2, right: (1 + Math.SQRT1_2) / 2, bottom: (1 + Math.SQRT1_2) / 2}
        : geometry === 'flowChartDisplay' ? {left: 1 / 6, top: 0, right: 5 / 6, bottom: 1}
        : geometry === 'flowChartOfflineStorage' ? {left: 0.25, top: 0, right: 0.75, bottom: 0.5}
        : geometry === 'flowChartOnlineStorage' ? {left: 1 / 6, top: 0, right: 5 / 6, bottom: 1}
        : geometry === 'flowChartPunchedTape' ? {left: 0, top: 1 / 5, right: 1, bottom: 4 / 5}
        : geometry === 'flowChartMagneticDisk' ? {left: 0, top: 1 / 3, right: 1, bottom: 5 / 6}
        : geometry === 'flowChartMagneticDrum' ? {left: 1 / 6, top: 0, right: 5 / 6, bottom: 1}
        : geometry === 'flowChartMagneticTape' ? {left: magneticTapeTextInset, top: magneticTapeTextInset, right: 1 - magneticTapeTextInset, bottom: 1 - magneticTapeTextInset}
        : geometry === 'flowChartMultidocument' ? {left: 0, top: 3675 / 21600, right: 18595 / 21600, bottom: 20782 / 21600}
        : geometry === 'flowChartInternalStorage' ? {left: 1 / 8, top: 1 / 8, right: 1, bottom: 1}
        : ['roundRect', 'flowChartAlternateProcess'].includes(geometry) ? {left: Math.min(transform.width, transform.height) * (presetAdjustment ?? 0.16667) * 0.29289 / transform.width, top: Math.min(transform.width, transform.height) * (presetAdjustment ?? 0.16667) * 0.29289 / transform.height, right: 1 - Math.min(transform.width, transform.height) * (presetAdjustment ?? 0.16667) * 0.29289 / transform.width, bottom: 1 - Math.min(transform.width, transform.height) * (presetAdjustment ?? 0.16667) * 0.29289 / transform.height}
        : geometry === 'rtTriangle' ? {left: 1 / 12, top: 7 / 12, right: 7 / 12, bottom: 11 / 12}
        : geometry === 'parallelogram' ? {left: 1 / 12 + (presetAdjustment ?? 0.25) * 5 * shortSide / transform.width / 12, top: 1 / 12 + (presetAdjustment ?? 0.25) * 5 * shortSide / transform.width / 12, right: 11 / 12 - (presetAdjustment ?? 0.25) * 5 * shortSide / transform.width / 12, bottom: 11 / 12 - (presetAdjustment ?? 0.25) * shortSide / transform.width / 12 * 5}
        : geometry === 'hexagon' ? hexagonTextRect : geometry === 'chevron' ? chevronTextRect
        : geometry === 'trapezoid' ? {left: trapezoidTextInset, top: trapezoidTextInset, right: 1 - trapezoidTextInset, bottom: 1}
        : geometry === 'octagon' ? octagonTextRect
        : geometry === 'pentagon' ? pentagonTextRect
        : geometry === 'star5' ? star5TextRect
        : geometry === 'star6' ? star6TextRect
        : geometry === 'star8' ? star8TextRect
        : geometry === 'star4' ? star4TextRect
        : geometry === 'star7' ? star7TextRect
        : geometry === 'star10' ? star10TextRect
        : geometry === 'star12' ? star12TextRect
        : ['star16', 'star24', 'star32'].includes(geometry) ? highPointStarTextRect
        : geometry === 'dodecagon' ? {left: dodecagonTextInset, top: dodecagonTextInset, right: 1 - dodecagonTextInset, bottom: 1 - dodecagonTextInset}
        : geometry === 'mathMinus' ? mathMinusTextRect
        : geometry === 'mathEqual' ? mathEqualTextRect
        : geometry === 'mathPlus' ? mathPlusTextRect
        : geometry === 'mathDivide' ? mathDivideTextRect
        : geometry === 'mathMultiply' ? mathMultiplyTextRect
        : geometry === 'mathNotEqual' ? mathNotEqualTextRect
        : geometry === 'round1Rect' ? round1RectTextRect
        : geometry === 'round2DiagRect' ? round2DiagRectTextRect
        : geometry === 'round2SameRect' ? round2SameRectTextRect
        : geometry === 'snipRoundRect' ? snipRoundRectTextRect
        : geometry === 'snip1Rect' ? snip1RectTextRect
        : geometry === 'snip2DiagRect' ? snip2DiagRectTextRect
        : geometry === 'snip2SameRect' ? snip2SameRectTextRect
        : geometry === 'homePlate' ? homePlateTextRect
        : geometry === 'cube' ? cubeTextRect
        : geometry === 'plaque' ? plaqueTextRect
        : geometry === 'donut' ? donutTextRect
        : ['pie', 'chord'].includes(geometry) ? donutTextRect
        : geometry === 'pieWedge' ? pieWedgeTextRect
        : ['arc', 'blockArc'].includes(geometry) ? angleArcTextRect
        : geometry === 'frame' ? frameTextRect
        : ['cornerTabs', 'plaqueTabs', 'squareTabs'].includes(geometry) ? tabTextRect
        : geometry === 'nonIsoscelesTrapezoid' ? nonIsoscelesTrapezoidTextRect
        : geometry === 'heptagon' ? heptagonTextRect
        : geometry === 'decagon' ? decagonTextRect
        : geometry === 'plus' ? plusTextRect
        : geometry === 'leftRightArrow' ? leftRightArrowTextRect
        : geometry === 'upDownArrow' ? upDownArrowTextRect
        : geometry === 'upArrow' ? upArrowTextRect
        : geometry === 'downArrow' ? downArrowTextRect
        : geometry === 'bentUpArrow' ? bentUpArrowTextRect
        : geometry === 'quadArrow' ? quadArrowTextRect
        : geometry === 'leftRightUpArrow' ? leftRightUpArrowTextRect
        : geometry === 'rightArrow' ? rightArrowTextRect
        : geometry === 'leftArrow' ? leftArrowTextRect
        : geometry === 'notchedRightArrow' ? notchedRightArrowTextRect : undefined
    const supportedGeometry = PRESET_GEOMETRIES_WITH_HANDWRITTEN_RENDERERS
    if (!supportedGeometry.has(geometry) && !LINE_GEOMETRIES.has(geometry) && !customPaths?.length) warnings.add(`The ${geometry} shape is simplified to a rectangle.`)
    if (geometry === 'custom' && !customPaths) warnings.add('An unsupported custom geometry was rendered as a rectangle.')
    const placeholder = parsePlaceholderIdentity(shape)
    if (placeholder && warnPlaceholder) warnings.add('Placeholder properties are only partially inherited from the layout or master.')
    const tx = parseText(shape, theme, warnings, inheritedStyle, inheritedParagraphStyles, rels, groupFill)
    const nonVisual = firstDescendant(shape, 'cNvPr')
    const resolvedLine = referencedShapeLine(shape, theme, warnings, fillStyles)
    const line = resolvedLine.line
    let fill = children(properties).some(node => ['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'].includes(node.localName))
        ? fillColor(properties, theme, warnings, groupFill, true) : referencedShapeFill(shape, theme, warnings, fillStyles, true)
    const strokePatternValue = fillColor(line, resolvedLine.theme, warnings, groupFill, true)
    const strokePattern = strokePatternValue && isPatternFillCss(strokePatternValue) ? strokePatternValue : undefined
    let stroke = strokePattern ? patternForegroundColor(strokePattern) : fillColor(line, resolvedLine.theme, warnings, groupFill)
    if (strokePattern) warnings.add('DrawingML pattern outlines use SVG paint servers on SVG-rendered shapes and connectors; CSS-rendered shape outlines retain a solid-color fallback.')
    const strokeDash = lineDashStyle(line, warnings)
    const strokeCap = lineCapStyle(line, warnings)
    const usesSvgPresetStroke = ['leftBracket', 'rightBracket', 'bracketPair', 'leftBrace', 'rightBrace', 'bracePair', 'mathDivide', 'mathMultiply', 'mathNotEqual', 'round1Rect', 'round2DiagRect', 'round2SameRect', 'snipRoundRect', 'cube', 'plaque', 'frame', 'halfFrame', 'cornerTabs', 'plaqueTabs', 'squareTabs', 'donut', 'pie', 'chord', 'arc', 'blockArc', 'pieWedge', 'cloud'].includes(geometry)
    if (line && numberAttr(line, 'w') === 0 && stroke !== 'transparent' && !isLineGeometry && !customGeometry && !customPaths && !usesSvgPresetStroke && !['line', 'rect', 'ellipse', 'roundRect', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDelay', 'flowChartDisplay', 'flowChartOfflineStorage', 'flowChartOnlineStorage', 'flowChartPunchedTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartMagneticTape', 'flowChartMultidocument', 'flowChartInternalStorage', 'flowChartSummingJunction', 'flowChartOr', 'plus', 'mathMinus', 'mathEqual', 'mathPlus', 'snip1Rect', 'snip2DiagRect', 'snip2SameRect', 'homePlate', 'nonIsoscelesTrapezoid', 'heptagon', 'decagon', 'leftRightArrow', 'upDownArrow', 'upArrow', 'downArrow', 'bentUpArrow', 'quadArrow', 'leftRightUpArrow', 'rightArrow', 'leftArrow', 'notchedRightArrow', 'heart', 'lightningBolt', 'triangle', 'rtTriangle', 'diamond', 'flowChartConnector', 'flowChartDecision', 'flowChartSort', 'flowChartCollate', 'flowChartPunchedCard', 'flowChartOffpageConnector', 'flowChartDocument', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartInputOutput', 'flowChartPreparation', 'flowChartTerminator', 'flowChartPredefinedProcess', 'flowChartExtract', 'flowChartMerge', 'parallelogram', 'trapezoid', 'hexagon', 'octagon', 'pentagon', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32', 'dodecagon', 'chevron'].includes(geometry)) {
        warnings.add('Zero-width preset shape outlines use CSS borders, whose device-pixel thickness may be rounded by the browser.')
    }
    if (Array.isArray(strokeDash) && !isLineGeometry && !customGeometry && !customPaths && !usesSvgPresetStroke && !['line', 'rect', 'ellipse', 'roundRect', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDelay', 'flowChartDisplay', 'flowChartOfflineStorage', 'flowChartOnlineStorage', 'flowChartPunchedTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartMagneticTape', 'flowChartMultidocument', 'flowChartInternalStorage', 'flowChartSummingJunction', 'flowChartOr', 'plus', 'mathMinus', 'mathEqual', 'mathPlus', 'snip1Rect', 'snip2DiagRect', 'snip2SameRect', 'homePlate', 'nonIsoscelesTrapezoid', 'heptagon', 'decagon', 'leftRightArrow', 'upDownArrow', 'upArrow', 'downArrow', 'bentUpArrow', 'quadArrow', 'leftRightUpArrow', 'rightArrow', 'leftArrow', 'notchedRightArrow', 'triangle', 'rtTriangle', 'diamond', 'flowChartConnector', 'flowChartDecision', 'flowChartSort', 'flowChartCollate', 'flowChartPunchedCard', 'flowChartOffpageConnector', 'flowChartDocument', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartInputOutput', 'flowChartPreparation', 'flowChartTerminator', 'flowChartPredefinedProcess', 'flowChartExtract', 'flowChartMerge'].includes(geometry)) warnings.add('Custom dash ratios on preset shape outlines are approximated with CSS dashes.')
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
    const shapeEffects = referencedShapeEffects(shape, theme, warnings, fillStyles)
    return {
        id: nonVisual?.getAttribute('id') || crypto.randomUUID(),
        name: nonVisual?.getAttribute('name') || 'Shape',
        hyperlink: parseHyperlink(child(nonVisual, 'hlinkClick'), rels, warnings),
        kind: isLineGeometry ? 'line' : 'shape',
        geometry,
        linePath,
        presetAdjustment,
        presetAdjustments: Object.keys(presetAdjustments).length ? presetAdjustments : undefined,
        customPaths,
        customTextRect,
        ...transform,
        fill,
        stroke,
        strokePattern,
        strokeWidth: numberAttr(line, 'w'),
        strokeDash,
        strokeCap,
        lineHeadEnd: parseLineEnd('headEnd'),
        lineTailEnd: parseLineEnd('tailEnd'),
        ...shapeEffects,
        ...tx,
        placeholder
    }
}

function parseTable(frame: Element, theme: Record<string, string>, warnings: Set<string>, rels?: Map<string, PptxRelationship>): PptxElement | undefined {
    const table = firstDescendant(frame, 'tbl')
    if (!table) return undefined
    const columns = children(child(table, 'tblGrid'), 'gridCol').map(column => numberAttr(column, 'w'))
    const rows = children(table, 'tr').map(row => ({
        height: numberAttr(row, 'h'),
        cells: children(row, 'tc').map((cell): PptxTableCell => {
            const properties = child(cell, 'tcPr')
            const text = parseText(cell, theme, warnings, {}, {}, rels)
            if (text.textAutoFit === 'shape') warnings.add('a:spAutoFit on table cells is not rendered; original cell and row sizes are retained.')
            const border = (name: string) => {
                const line = child(properties, name)
                if (!line) return undefined
                return {color: fillColor(line, theme, warnings), width: numberAttr(line, 'w', 12_700)}
            }
            const anchor = properties?.getAttribute('anchor')
            return {
                paragraphs: text.paragraphs,
                textAnchorCenter: text.textAnchorCenter,
                textColumnCount: text.textColumnCount,
                textColumnSpacing: text.textColumnSpacing,
                textColumnsRightToLeft: text.textColumnsRightToLeft,
                textFontScale: text.textFontScale,
                textLineSpacingReduction: text.textLineSpacingReduction,
                textAutoFitDynamic: text.textAutoFitDynamic,
                fill: fillColor(properties, theme, warnings, undefined, true),
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
        hyperlink: parseHyperlink(child(nonVisual, 'hlinkClick'), rels, warnings),
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

async function parseSmartArtCachedDrawing(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>, frameAnimationTargetIds: Set<string>, parentTransform: Matrix2D | undefined, sourceTheme: Record<string, string>, fillStyles: ParsedTheme | undefined, smartArtAnimationTargets: Map<string, string>): Promise<PptxElement[] | undefined> {
    const graphicData = firstDescendant(frame, 'graphicData')
    if (!graphicData?.getAttribute('uri')?.endsWith('/diagram')) return undefined
    const frameId = firstDescendant(frame, 'cNvPr')?.getAttribute('id') || ''
    if (frameAnimationTargetIds.has(frameId)) {
        warnings.add('SmartArt with frame-targeted animation uses the text fallback to preserve its animation.')
        return undefined
    }
    const dataRelation = rels.get(namespacedAttr(child(graphicData, 'relIds'), 'dm') || '')
    const dataPart = dataRelation?.target.match(/^(.*\/)data(\d+)\.xml$/i)
    if (!dataRelation || dataRelation.external || !dataPart) return undefined
    const expectedDrawingPath = `${dataPart[1]}drawing${dataPart[2]}.xml`
    const drawingRelation = [...rels.values()].find(relation => !relation.external && relation.type.endsWith('/diagramDrawing') && relation.target.toLowerCase() === expectedDrawingPath.toLowerCase())
    if (!drawingRelation) return undefined
    const part = zip.file(drawingRelation.target)
    if (!part) {
        warnings.add('The SmartArt cached drawing part is missing; its text fallback was used.')
        return undefined
    }

    try {
        const xml = parseXml(await part.async('string'), drawingRelation.target)
        const tree = firstDescendant(xml.documentElement, 'spTree')
        const groupTransform = tree && groupTransformMatrix(tree)
        const groupXfrm = tree && child(child(tree, 'grpSpPr'), 'xfrm')
        const groupOffset = child(groupXfrm, 'off')
        const groupExtent = child(groupXfrm, 'ext')
        const groupWidth = numberAttr(groupExtent, 'cx')
        const groupHeight = numberAttr(groupExtent, 'cy')
        const frameTransform = shapeTransform(frame)
        if (!tree || !groupTransform || !groupXfrm || groupWidth <= 0 || groupHeight <= 0 || frameTransform.width <= 0 || frameTransform.height <= 0) {
            warnings.add('The SmartArt cached drawing has invalid bounds; its text fallback was used.')
            return undefined
        }
        const framePlacement = multiplyMatrix(
            translationMatrix(frameTransform.x, frameTransform.y),
            multiplyMatrix(
                translationMatrix(frameTransform.width / 2, frameTransform.height / 2),
                multiplyMatrix(
                    rotationMatrix(frameTransform.rotation),
                    multiplyMatrix(
                        [frameTransform.flipH ? -1 : 1, 0, 0, frameTransform.flipV ? -1 : 1, 0, 0],
                        translationMatrix(-frameTransform.width / 2, -frameTransform.height / 2)
                    )
                )
            )
        )
        const drawingToFrame = multiplyMatrix(
            scaleMatrix(frameTransform.width / groupWidth, frameTransform.height / groupHeight),
            translationMatrix(-numberAttr(groupOffset, 'x'), -numberAttr(groupOffset, 'y'))
        )
        const drawingTransform = multiplyMatrix(parentTransform || [1, 0, 0, 1, 0, 0], multiplyMatrix(framePlacement, multiplyMatrix(drawingToFrame, groupTransform)))
        const drawingRels = await loadRelationships(zip, relationshipPath(drawingRelation.target))
        const elements: PptxElement[] = []
        for (const node of children(tree)) {
            elements.push(...(await parseSceneElements(node, drawingRels, zip, theme, warnings, urls, imageUrlCache, drawingTransform, frameAnimationTargetIds, sourceTheme, fillStyles, undefined, smartArtAnimationTargets)))
        }
        if (!elements.length) return undefined
        warnings.add('SmartArt uses its last saved diagram drawing; its data-driven layout is not regenerated in the browser.')
        const frameName = firstDescendant(frame, 'cNvPr')?.getAttribute('name') || 'SmartArt'
        return elements.map((element, index) => {
            const sourceId = element.id || String(index)
            const id = `${frameId}:smartart:${drawingRelation.target}:${sourceId}`
            if (element.id) smartArtAnimationTargets.set(`${frameId}:${element.id}`, id)
            return {...element, id, name: `${frameName}: ${element.name}`}
        })
    } catch {
        warnings.add('The SmartArt cached drawing could not be parsed; its text fallback was used.')
        return undefined
    }
}

async function parseSmartArtTextFallback(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>): Promise<PptxElement | undefined> {
    const graphicData = firstDescendant(frame, 'graphicData')
    if (!graphicData?.getAttribute('uri')?.endsWith('/diagram')) return undefined

    const dataId = namespacedAttr(child(graphicData, 'relIds'), 'dm') || ''
    const relation = rels.get(dataId)
    let paragraphs: PptxParagraph[] = []
    let hasHierarchy = false
    if (!relation || relation.external || !relation.type.endsWith('/diagramData')) {
        warnings.add('SmartArt diagram data is missing or external; its text could not be recovered.')
    } else {
        const part = zip.file(relation.target)
        if (!part) warnings.add('The SmartArt diagram data part is missing; its text could not be recovered.')
        else {
            try {
                const xml = parseXml(await part.async('string'), relation.target)
                const allPoints = children(firstDescendant(xml.documentElement, 'ptLst'), 'pt')
                const points = allPoints.filter(point => point.getAttribute('type') === 'node')
                const pointIds = new Set(allPoints.map(point => point.getAttribute('modelId')).filter((id): id is string => Boolean(id)))
                const pointTypes = new Map(allPoints.map(point => [point.getAttribute('modelId') || '', point.getAttribute('type') || '']))
                const parentByChild = new Map<string, string>()
                for (const connection of children(firstDescendant(xml.documentElement, 'cxnLst'), 'cxn')) {
                    if (!['parOf', 'presParOf'].includes(connection.getAttribute('type') || '')) continue
                    const parent = connection.getAttribute('srcId') || ''
                    const node = connection.getAttribute('destId') || ''
                    if (!pointIds.has(parent) || !pointIds.has(node) || parent === node) {
                        warnings.add('A SmartArt parent connection with an invalid endpoint was ignored.')
                        continue
                    }
                    const existingParent = parentByChild.get(node)
                    if (existingParent && existingParent !== parent) {
                        warnings.add('A SmartArt node with multiple parents uses its first parent connection.')
                        continue
                    }
                    parentByChild.set(node, parent)
                }
                const depthByNode = new Map<string, number>()
                let warnedCycle = false
                let warnedDepth = false
                for (const point of points) {
                    const id = point.getAttribute('modelId') || ''
                    const visited = new Set([id])
                    let current = id
                    let depth = 0
                    while (parentByChild.has(current)) {
                        const parent = parentByChild.get(current)!
                        if (pointTypes.get(parent) === 'doc') break
                        if (visited.has(parent)) {
                            if (!warnedCycle) warnings.add('A cyclic SmartArt parent connection was ignored while computing indentation.')
                            warnedCycle = true
                            depth = 0
                            break
                        }
                        visited.add(parent)
                        current = parent
                        depth++
                        if (depth >= 8 && parentByChild.has(current)) {
                            if (!warnedDepth) warnings.add('SmartArt hierarchy indentation was capped at eight levels.')
                            warnedDepth = true
                            depth = 8
                            break
                        }
                    }
                    depthByNode.set(id, depth)
                }
                hasHierarchy = [...depthByNode.values()].some(depth => depth > 0)
                for (const point of points) {
                    const pointText = child(point, 't')
                    if (!pointText) continue
                    const shape = pointText.ownerDocument.createElementNS(PRESENTATION_NS, 'p:sp')
                    const body = pointText.ownerDocument.createElementNS(PRESENTATION_NS, 'p:txBody')
                    for (const node of Array.from(pointText.childNodes)) body.appendChild(node.cloneNode(true))
                    shape.appendChild(body)
                    const depth = depthByNode.get(point.getAttribute('modelId') || '') || 0
                    const indentation = depth * 304_800
                    paragraphs.push(...parseText(shape, theme, warnings).paragraphs.map(paragraph => ({
                        ...paragraph,
                        marginLeft: (paragraph.marginLeft || 0) + indentation
                    })))
                }
            } catch {
                warnings.add('The SmartArt diagram data contains invalid XML; its text could not be recovered.')
            }
        }
    }

    if (!paragraphs.some(paragraph => paragraph.runs.some(run => run.text.trim()))) {
        paragraphs = [{align: 'left', runs: [{text: 'SmartArt text unavailable', fontSizePt: 14}]}]
    }
    warnings.add(hasHierarchy
        ? 'SmartArt is shown as an indented text list from its parent connections; node layout, connectors, and formatting are not reproduced.'
        : 'SmartArt is shown as a plain text list; node layout, hierarchy, connectors, and formatting are not reproduced.')
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

async function parseOlePreview(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>, fillStyles?: ParsedTheme): Promise<PptxElement | undefined> {
    const graphicData = firstDescendant(frame, 'graphicData')
    const ole = child(graphicData, 'oleObj')
    if (!ole && graphicData?.getAttribute('uri') !== OLE_GRAPHIC_DATA_URI) return undefined

    const preview = child(ole, 'pic')
    const previewBlip = preview && firstDescendant(preview, 'blip')
    const previewRelation = rels.get(namespacedAttr(previewBlip, 'embed') || '')
    if (preview && previewRelation && !previewRelation.external && zip.file(previewRelation.target)) {
        const picture = await parsePicture(preview, rels, zip, theme, warnings, urls, imageUrlCache, fillStyles)
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
        const rawValue = child(point, 'v')?.textContent ?? ''
        const value = rawValue.trim() ? rawValue : null
        return Number.isSafeInteger(index) && index >= 0 ? [{index, value}] : []
    })
    const declaredCount = numberAttr(child(cache, 'ptCount'), 'val', points.length)
    const count = Math.min(128, Math.max(declaredCount, ...points.map(point => point.index + 1), 0))
    if (declaredCount > 128 || points.some(point => point.index >= 128)) warnings.add('Chart data points beyond 128 are omitted.')
    const values: Array<string | null> = Array(count).fill(null)
    for (const point of points) {
        if (point.index < count) values[point.index] = point.value
    }
    return values
}

function chartDataValues(source: Element | undefined, workbook: ChartWorkbookMetadata | undefined, warnings: Set<string>): Array<string | null> {
    const hasCache = ['strCache', 'numCache', 'strLit', 'numLit'].some(name => !!firstDescendant(source, name))
    const cachedValues = hasCache ? chartCache(source, warnings) : []
    if (cachedValues.some(value => value !== null)) return cachedValues
    const reference = firstDescendant(source, 'numRef') || firstDescendant(source, 'strRef')
    const formula = child(reference, 'f')?.textContent?.trim()
    if (!formula) return cachedValues
    const values = workbook?.valuesByFormula.get(formula.toLowerCase())
    if (values?.some(value => value !== null)) {
        warnings.add('Chart references without saved data caches use values from the embedded workbook.')
        return values
    }
    warnings.add('A chart reference without a saved data cache could not be resolved from the embedded workbook; its values are omitted.')
    return []
}

function chartCachedNumberFormat(source: Element | undefined, warnings: Set<string>): string | undefined {
    const cache = firstDescendant(source, 'numCache')
    const formatCode = child(cache, 'formatCode')?.textContent?.trim()
    if (!formatCode || formatCode.toLowerCase() === 'general') return undefined
    if (isSupportedChartNumberFormat(formatCode)) return formatCode
    warnings.add('An unsupported numeric-cache chart format was ignored; an embedded workbook format may be used as fallback.')
    return undefined
}

function sharedChartNumberFormat(formats: Array<string | undefined>): string | undefined {
    return formats.length && formats.every(format => format && format === formats[0]) ? formats[0] : undefined
}

function chartWorkbookNumberFormat(source: Element | undefined, workbook: ChartWorkbookMetadata | undefined, warnings: Set<string>): string | undefined {
    const formula = child(firstDescendant(source, 'numRef'), 'f')?.textContent?.trim()
    if (!formula || !workbook) return undefined
    const formats = workbook.numberFormatsByFormula.get(formula.toLowerCase())
    if (!formats?.length) return undefined
    const format = sharedChartNumberFormat(formats)
    if (format) return format
    if (formats.some(Boolean)) warnings.add('Source-linked chart references use mixed embedded workbook formats and are rendered as raw values.')
    return undefined
}

function chartFillColor(properties: Element | undefined, theme: Record<string, string>, warnings: Set<string>, fallback: string): string {
    const color = fillColor(properties, theme, warnings, undefined, true)
    if (isPatternFillCss(color)) return color
    if (color.includes('gradient')) {
        const firstStop = firstDescendant(child(properties, 'gradFill'), 'gs')
        warnings.add('Chart gradient fills are approximated with their first color stop.')
        return readColor(children(firstStop)[0], theme) || fallback
    }
    return color === 'transparent' ? fallback : color
}

function chartStrokeWidthEmu(line: Element | undefined, warnings: Set<string>): number | undefined {
    const raw = line?.getAttribute('w')
    if (raw === null || raw === undefined) return undefined
    const width = Number(raw)
    if (!/^\d+$/.test(raw.trim()) || !Number.isSafeInteger(width) || width > MAX_ANIMATED_STROKE_WIDTH_EMU) {
        warnings.add('An invalid chart line width was ignored.')
        return undefined
    }
    return width
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

function chartDataLabelTextStyle(dataLabels: Element | undefined, theme: Record<string, string>, warnings: Set<string>): PptxChartDataLabelStyle | undefined {
    const textProperties = child(dataLabels, 'txPr')
    if (!textProperties) return undefined
    const runProperties = firstDescendant(textProperties, 'defRPr') || firstDescendant(textProperties, 'endParaRPr') || firstDescendant(textProperties, 'rPr')
    if (!runProperties) {
        warnings.add('Chart data label paragraph text styles are not rendered.')
        return undefined
    }
    const parsed = resolveTextFonts(textStyle(runProperties, theme, warnings), '', theme, warnings)
    if (
        (parsed.underline && parsed.lineThrough && (
            parsed.underlineStyle !== parsed.lineThroughStyle || parsed.underlineColor !== undefined || parsed.underlineThicknessEmu !== undefined
        )) ||
        parsed.fillPattern || parsed.fillGradient || parsed.highlightColor || (parsed.outlineWidth !== undefined && parsed.outlineWidth > 0) || parsed.shadow || parsed.shadows?.length || parsed.glow || parsed.glows?.length
    ) warnings.add('Advanced chart data label decorations, outlines, and effects are simplified.')
    if (parsed.fontFamilyEastAsia || parsed.fontFamilyComplexScript) warnings.add('East Asian and complex-script chart label fonts use browser fallback.')
    const paragraphProperties = firstDescendant(textProperties, 'pPr')
    const bodyProperties = firstDescendant(textProperties, 'bodyPr')
    if (
        (paragraphProperties && (paragraphProperties.attributes.length > 0 || children(paragraphProperties).some(node => node.localName !== 'defRPr'))) ||
        (bodyProperties && (bodyProperties.attributes.length > 0 || children(bodyProperties).length > 0))
    ) warnings.add('Chart data label paragraph layout and text orientation are simplified.')
    const style: PptxChartDataLabelStyle = {}
    if (parsed.fontFamily !== undefined) style.fontFamily = parsed.fontFamily
    if (parsed.fontSizePt !== undefined) style.fontSizePt = parsed.fontSizePt
    if (parsed.color !== undefined) style.color = parsed.color
    if (parsed.bold !== undefined) style.bold = parsed.bold
    if (parsed.italic !== undefined) style.italic = parsed.italic
    if (parsed.letterSpacingPt !== undefined) style.letterSpacingPt = parsed.letterSpacingPt
    if (parsed.baselineOffsetEm !== undefined) style.baselineOffsetEm = parsed.baselineOffsetEm
    if (parsed.underline !== undefined) style.underline = parsed.underline
    if (parsed.underlineStyle !== undefined) style.underlineStyle = parsed.underlineStyle
    if (parsed.underlineColor !== undefined) style.underlineColor = parsed.underlineColor
    if (parsed.underlineThicknessEmu !== undefined) style.underlineThicknessEmu = parsed.underlineThicknessEmu
    if (parsed.lineThrough !== undefined) style.lineThrough = parsed.lineThrough
    if (parsed.lineThroughStyle !== undefined) style.lineThroughStyle = parsed.lineThroughStyle
    return Object.keys(style).length ? style : undefined
}

function chartDataLabelTextStyleOverrides(dataLabels: Element | undefined, theme: Record<string, string>, warnings: Set<string>): Record<number, PptxChartDataLabelStyle> {
    const overrides: Record<number, PptxChartDataLabelStyle> = {}
    for (const label of children(dataLabels, 'dLbl')) {
        const index = numberAttr(child(label, 'idx'), 'val', -1)
        const style = chartDataLabelTextStyle(label, theme, warnings)
        if (index >= 0 && style) overrides[index] = style
    }
    return overrides
}

function mergeChartDataLabelTextStyleOverrides(...sources: Array<Record<number, PptxChartDataLabelStyle>>): Record<number, PptxChartDataLabelStyle> {
    const result: Record<number, PptxChartDataLabelStyle> = {}
    for (const source of sources) {
        for (const [index, style] of Object.entries(source)) result[Number(index)] = {...result[Number(index)], ...style}
    }
    return result
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

const chartMarkerSymbols = new Set<PptxChartMarkerSymbol>(['auto', 'circle', 'dash', 'diamond', 'dot', 'none', 'picture', 'plus', 'square', 'star', 'triangle', 'x'])

function chartMarkerSymbol(marker: Element | undefined, warnings: Set<string>): PptxChartMarkerSymbol | undefined {
    const symbolNode = child(marker, 'symbol')
    if (!symbolNode) return undefined
    const value = symbolNode.getAttribute('val')?.trim() || 'auto'
    if (!chartMarkerSymbols.has(value as PptxChartMarkerSymbol)) {
        warnings.add(`Unsupported chart marker symbol "${value}" is rendered as a circle.`)
        return 'circle'
    }
    if (value === 'picture') {
        warnings.add('Picture chart markers are rendered as circles.')
        return 'circle'
    }
    return value as PptxChartMarkerSymbol
}

function chartMarkerSize(marker: Element | undefined, warnings: Set<string>): number | undefined {
    const sizeNode = child(marker, 'size')
    if (!sizeNode) return undefined
    const value = Number(sizeNode.getAttribute('val'))
    if (!Number.isInteger(value) || value < 2 || value > 72) {
        warnings.add('An invalid chart marker size uses the default size.')
        return undefined
    }
    return value
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

function chartAxisTickMark(axis: Element | undefined, axisName: string, kind: 'major' | 'minor', warnings: Set<string>): PptxChartTickMark | undefined {
    const node = child(axis, `${kind}TickMark`)
    if (!node) return undefined
    const value = node.getAttribute('val')
    const tickMark = value === null ? 'cross' : value
    if (['cross', 'in', 'none', 'out'].includes(tickMark)) return tickMark as PptxChartTickMark
    warnings.add(`Unsupported chart ${axisName} ${kind} tick mark is ignored.`)
    return undefined
}

function chartAxisSkip(axis: Element | undefined, name: 'tickLblSkip' | 'tickMarkSkip', warnings: Set<string>): number | undefined {
    const node = child(axis, name)
    if (!node) return undefined
    const rawValue = node.getAttribute('val')?.trim() || ''
    const value = Number(rawValue)
    if (!/^\d+$/.test(rawValue) || !Number.isSafeInteger(value) || value < 1) {
        warnings.add(`Invalid chart category-axis ${name} value is ignored.`)
        return undefined
    }
    return value
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

function chartValueAxis(axis: Element | undefined, warnings: Set<string>, axisName = 'value', sourceNumberFormat?: string): (Omit<PptxChartAxis, 'min' | 'max'> & {min?: number; max?: number}) | undefined {
    if (!axis) return undefined
    const deleted = chartAxisDeleted(axis, axisName, warnings)
    const scaling = child(axis, 'scaling')
    const logBaseNode = child(scaling, 'logBase')
    const parsedLogBase = logBaseNode?.getAttribute('val')?.trim() || ''
    const logBaseValue = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(parsedLogBase) ? Number(parsedLogBase) : NaN
    const logBase = Number.isFinite(logBaseValue) && logBaseValue >= 2 && logBaseValue <= 1000 ? logBaseValue : undefined
    if (logBaseNode && logBase === undefined) warnings.add(`Invalid chart ${axisName}-axis logarithmic base is ignored; c:logBase must be between 2 and 1000.`)
    const readBound = (name: 'min' | 'max') => {
        const node = child(scaling, name)
        if (!node) return undefined
        const raw = node.getAttribute('val')
        const value = raw?.trim() ? Number(raw) : NaN
        if (!Number.isFinite(value)) {
            warnings.add(`Invalid chart ${axisName}-axis bounds are ignored.`)
            return undefined
        }
        return value
    }
    let min = readBound('min')
    let max = readBound('max')
    if (min !== undefined && max !== undefined && min >= max) {
        warnings.add(`Invalid chart ${axisName}-axis bounds are ignored.`)
        min = undefined
        max = undefined
    }
    const majorUnitNode = child(axis, 'majorUnit')
    const majorUnit = majorUnitNode ? Number(majorUnitNode.getAttribute('val')) : undefined
    if (majorUnitNode && (!Number.isFinite(majorUnit) || majorUnit! <= 0)) warnings.add(`Invalid chart ${axisName}-axis major units are ignored.`)
    const minorUnitNode = child(axis, 'minorUnit')
    const minorUnit = minorUnitNode ? Number(minorUnitNode.getAttribute('val')) : undefined
    if (minorUnitNode && (!Number.isFinite(minorUnit) || minorUnit! <= 0)) warnings.add(`Invalid chart ${axisName}-axis minor units are ignored.`)
    const numberFormatNode = child(axis, 'numFmt')
    const sourceLinked = (numberFormatNode?.getAttribute('sourceLinked') || '1').toLowerCase()
    const formatCode = numberFormatNode?.getAttribute('formatCode')?.trim()
    const position = chartAxisPosition(axis, axisName, warnings)
    const crossing = chartAxisCrossing(axis, axisName, warnings)
    const majorTickMark = chartAxisTickMark(axis, axisName, 'major', warnings)
    const minorTickMark = chartAxisTickMark(axis, axisName, 'minor', warnings)
    if (minorTickMark && minorTickMark !== 'none' && !(Number.isFinite(minorUnit) && minorUnit! > 0)) warnings.add(logBase !== undefined
        ? `Automatic chart ${axisName}-axis logarithmic minor intervals use base subdivisions.`
        : `Automatic chart ${axisName}-axis minor intervals use four subdivisions per major interval.`)
    const tickLabelPosition = chartTickLabelPosition(axis, axisName, warnings)
    let numberFormat: string | undefined
    if (formatCode && ['0', 'false'].includes(sourceLinked)) {
        if (isSupportedChartNumberFormat(formatCode)) numberFormat = formatCode
        else warnings.add('Unsupported chart number formats are rendered as raw values.')
    } else if (formatCode) {
        if (sourceNumberFormat) numberFormat = sourceNumberFormat
        else warnings.add('Source-linked chart number formats use raw values when no supported cached or embedded workbook style is available.')
    }
    return {
        ...(min !== undefined ? {min} : {}),
        ...(max !== undefined ? {max} : {}),
        ...(logBase !== undefined ? {logBase} : {}),
        ...(Number.isFinite(majorUnit) && majorUnit! > 0 ? {majorUnit} : {}),
        ...(Number.isFinite(minorUnit) && minorUnit! > 0 ? {minorUnit} : {}),
        ...(numberFormat ? {numberFormat} : {}),
        ...(child(scaling, 'orientation')?.getAttribute('val') === 'maxMin' ? {reverse: true} : {}),
        ...(deleted ? {deleted: true} : {}),
        ...(child(axis, 'majorGridlines') ? {majorGridlines: true} : {}),
        ...(position ? {position} : {}),
        ...(majorTickMark ? {majorTickMark} : {}),
        ...(minorTickMark ? {minorTickMark} : {}),
        ...(tickLabelPosition ? {tickLabelPosition} : {}),
        ...crossing
    }
}

async function parseChart(frame: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, sourceTheme: Record<string, string>): Promise<PptxElement | undefined> {
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
    const chartDate1904 = chartBoolean(child(xml.documentElement, 'date1904')) ?? false
    const chartWorkbook = await chartWorkbookMetadata(xml.documentElement, relation.target, zip, warnings)
    const indexedColors = chartWorkbook?.indexedColors
    // Chart clrMapOvr carries CT_ColorMapping attributes directly on the element.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.colormapoverride
    const colorMap = child(xml.documentElement, 'clrMapOvr')
    if (colorMap) theme = mappedThemeColors(sourceTheme, readColorMap(colorMap, {}, warnings), theme)
    const chart = firstDescendant(xml.documentElement, 'chart')
    const plotArea = firstDescendant(chart, 'plotArea')
    const supported = ['barChart', 'lineChart', 'areaChart', 'pieChart', 'doughnutChart', 'scatterChart', 'bubbleChart', 'radarChart', 'stockChart']
    const chartNodes = children(plotArea).filter(item => item.localName.endsWith('Chart'))
    const chartAxisIds = (node: Element) => children(node, 'axId').map(item => item.getAttribute('val') || '').filter(Boolean)
    const stockChartNode = chartNodes.find(item => item.localName === 'stockChart')
    const volumeChartNode = chartNodes.find(item => item.localName === 'barChart')
    const stockAxisIds = stockChartNode ? chartAxisIds(stockChartNode) : []
    const volumeAxisIds = volumeChartNode ? chartAxisIds(volumeChartNode) : []
    const stockVolumeMode = chartNodes.length === 2
        && !!stockChartNode
        && !!volumeChartNode
        && children(volumeChartNode, 'ser').length === 1
        && child(volumeChartNode, 'barDir')?.getAttribute('val') !== 'bar'
        && stockAxisIds.length >= 2
        && volumeAxisIds.length >= 2
        && stockAxisIds[0] === volumeAxisIds[0]
        && stockAxisIds[1] !== volumeAxisIds[1]
    const typeNode = stockVolumeMode ? stockChartNode! : chartNodes.find(item => supported.includes(item.localName))
    if (!typeNode) {
        warnings.add(`The ${chartNodes[0]?.localName || 'unknown'} chart type is not rendered.`)
        return undefined
    }
    const comboMode = stockVolumeMode || (chartNodes.length > 1 && chartNodes.every(item => ['barChart', 'lineChart', 'areaChart'].includes(item.localName)) && new Set(chartNodes.map(item => item.localName)).size > 1)
    if (chartNodes.length > 1 && !comboMode) warnings.add('A mixed chart is simplified to its first supported chart type.')
    const primaryGroupAxisIds = chartAxisIds(typeNode)
    const primaryValueAxisId = primaryGroupAxisIds[1]
    const type = typeNode.localName === 'barChart' ? 'bar' : typeNode.localName === 'lineChart' ? 'line' : typeNode.localName === 'areaChart' ? 'area' : typeNode.localName === 'doughnutChart' ? 'doughnut' : typeNode.localName === 'scatterChart' ? 'scatter' : typeNode.localName === 'bubbleChart' ? 'bubble' : typeNode.localName === 'radarChart' ? 'radar' : typeNode.localName === 'stockChart' ? 'stock' : 'pie'
    const isPieLike = type === 'pie' || type === 'doughnut'
    const isScatter = type === 'scatter'
    const isBubble = type === 'bubble'
    const isRadar = type === 'radar'
    const isStock = type === 'stock'
    const isXYChart = isScatter || isBubble
    const radarStyleValue = isRadar ? child(typeNode, 'radarStyle')?.getAttribute('val') || 'standard' : undefined
    const radarStyle = radarStyleValue === 'marker' || radarStyleValue === 'filled' ? radarStyleValue : 'standard'
    if (isRadar && radarStyle !== radarStyleValue) warnings.add('Unsupported radar style is rendered as standard lines.')
    const scatterStyleValue = child(typeNode, 'scatterStyle')?.getAttribute('val') || 'lineMarker'
    const scatterStyle = ['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker'].includes(scatterStyleValue) ? scatterStyleValue : 'marker'
    if (isScatter && scatterStyle !== scatterStyleValue) warnings.add('Unsupported scatter style is rendered as markers.')
    const scatterShowLine = isScatter && ['line', 'lineMarker', 'smooth', 'smoothMarker'].includes(scatterStyle)
    const scatterShowMarker = isScatter && ['marker', 'lineMarker', 'smoothMarker'].includes(scatterStyle)
    const firstSliceAngleNode = child(typeNode, 'firstSliceAng')
    const firstSliceAngleValue = firstSliceAngleNode ? Number(firstSliceAngleNode.getAttribute('val')) : 0
    const firstSliceAngle = Number.isInteger(firstSliceAngleValue) && firstSliceAngleValue >= 0 && firstSliceAngleValue <= 360 ? firstSliceAngleValue : 0
    if (firstSliceAngleNode && firstSliceAngle !== firstSliceAngleValue) warnings.add('Invalid chart first-slice angle uses 0 degrees.')
    const holeSizeNode = type === 'doughnut' ? child(typeNode, 'holeSize') : undefined
    const holeSizeValue = holeSizeNode ? Number(holeSizeNode.getAttribute('val')) : 50
    const holeSize = Number.isInteger(holeSizeValue) && holeSizeValue >= 10 && holeSizeValue <= 90 ? holeSizeValue : 50
    if (holeSizeNode && holeSize !== holeSizeValue) warnings.add('Invalid doughnut hole size uses the 50 percent default.')
    const bubbleScaleNode = isBubble ? child(typeNode, 'bubbleScale') : undefined
    const bubbleScaleValue = bubbleScaleNode?.getAttribute('val') == null ? 100 : Number(bubbleScaleNode.getAttribute('val'))
    const bubbleScale = Number.isInteger(bubbleScaleValue) && bubbleScaleValue >= 0 && bubbleScaleValue <= 300 ? bubbleScaleValue : 100
    if (bubbleScaleNode && bubbleScale !== bubbleScaleValue) warnings.add('Invalid bubble scale uses the 100 percent default.')
    const bubbleSizeRepresentsValue = isBubble ? child(typeNode, 'sizeRepresents')?.getAttribute('val') || 'area' : undefined
    const bubbleSizeRepresents = bubbleSizeRepresentsValue === 'w' ? 'w' : 'area'
    if (isBubble && bubbleSizeRepresentsValue !== 'area' && bubbleSizeRepresentsValue !== 'w') warnings.add('Unsupported bubble size representation uses area.')
    const showNegativeBubbles = isBubble ? chartBoolean(child(typeNode, 'showNegBubbles')) ?? true : undefined
    const bubbleVaryColors = isBubble ? chartBoolean(child(typeNode, 'varyColors')) ?? false : undefined
    if (isBubble && chartBoolean(child(typeNode, 'bubble3D')) === true) warnings.add('Three-dimensional bubbles are approximated as flat circles.')
    const directionNode = comboMode ? chartNodes.find(item => item.localName === 'barChart')! : typeNode
    const direction = type === 'bar' || comboMode ? child(directionNode, 'barDir')?.getAttribute('val') === 'bar' ? 'horizontal' : 'vertical' : 'vertical'
    const groupingValue = child(typeNode, 'grouping')?.getAttribute('val') || 'standard'
    const stackedGrouping = groupingValue === 'stacked' || groupingValue === 'percentStacked'
    const grouping: PptxChartGrouping = ['bar', 'line', 'area'].includes(type) && stackedGrouping ? groupingValue : 'clustered'
    if (!['standard', 'clustered', 'stacked', 'percentStacked'].includes(groupingValue) || (!['bar', 'line', 'area'].includes(type) && stackedGrouping)) {
        warnings.add('Unsupported chart grouping is rendered as grouped series.')
    }
    if (!isPieLike && children(plotArea).some(item => ['catAx', 'dateAx', 'serAx'].includes(item.localName))) warnings.add('Chart date and series axes, minor tick marks, and advanced styling are simplified.')
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
    const seriesDescriptors = (comboMode ? chartNodes : [typeNode]).flatMap((groupNode, chartGroupIndex) => children(groupNode, 'ser').map(node => ({
        node,
        groupNode,
        chartGroupIndex,
        chartType: groupNode.localName === 'lineChart' ? 'line' as const : groupNode.localName === 'areaChart' ? 'area' as const : groupNode.localName === 'stockChart' ? 'stock' as const : 'bar' as const,
        grouping: child(groupNode, 'grouping')?.getAttribute('val') === 'stacked' ? 'stacked' as const : child(groupNode, 'grouping')?.getAttribute('val') === 'percentStacked' ? 'percentStacked' as const : 'clustered' as const,
        valueAxisId: chartAxisIds(groupNode)[1] || primaryValueAxisId
    })))
    const seriesNodes = seriesDescriptors.map(item => item.node)
    const parseNumericValues = (source: Element | undefined) => chartDataValues(source, chartWorkbook, warnings).map(value => {
        if (value === null || value.trim() === '') return null
        const number = Number(value)
        if (!Number.isFinite(number)) warnings.add('Malformed chart data points are skipped.')
        return Number.isFinite(number) ? number : null
    })
    const chartDataLabels = child(typeNode, 'dLbls')
    const chartShowLegendKey = chartDataLabelFlag(chartDataLabels, 'showLegendKey') ?? false
    const chartShowValueLabels = chartDataLabelFlag(chartDataLabels, 'showVal') ?? false
    const chartShowPercentLabels = chartDataLabelFlag(chartDataLabels, 'showPercent') ?? false
    const chartShowCategoryNameLabels = chartDataLabelFlag(chartDataLabels, 'showCatName') ?? false
    const chartShowSeriesNameLabels = chartDataLabelFlag(chartDataLabels, 'showSerName') ?? false
    const chartShowBubbleSizeLabels = chartDataLabelFlag(chartDataLabels, 'showBubbleSize') ?? false
    const chartShowLeaderLines = chartDataLabelFlag(chartDataLabels, 'showLeaderLines') ?? false
    const chartLabelSeparator = chartDataLabelSeparator(chartDataLabels)
    const chartValueFormat = chartNumberFormat(chartDataLabels)
    const chartLabelPosition = chartDataLabelPosition(chartDataLabels)
    const dataLabelNodes = [...new Set([...(comboMode ? chartNodes.map(item => child(item, 'dLbls')) : [chartDataLabels]), ...seriesNodes.map(item => child(item, 'dLbls'))].filter((item): item is Element => !!item))]
    const pointDataLabels = dataLabelNodes.flatMap(item => children(item, 'dLbl'))
    const allDataLabelSettings = [...dataLabelNodes, ...pointDataLabels]
    const numberFormatNodes = allDataLabelSettings.map(item => child(item, 'numFmt')).filter((item): item is Element => !!item)
    if (numberFormatNodes.some(item => ['0', 'false'].includes((item.getAttribute('sourceLinked') || '1').toLowerCase()) && item.getAttribute('formatCode') && !isSupportedChartNumberFormat(item.getAttribute('formatCode')!.trim()))) {
        warnings.add('Unsupported chart number formats are rendered as raw values.')
    }
    if (!isPieLike && allDataLabelSettings.some(item => chartBoolean(child(item, 'showPercent')) === true)) {
        warnings.add('Percentage data labels are rendered only for pie and doughnut charts.')
    }
    if (!isPieLike && allDataLabelSettings.some(item => chartDataLabelFlag(item, 'showLeaderLines') === true)) {
        warnings.add('Chart leader lines are rendered only for pie and doughnut charts.')
    }
    if (dataLabelNodes.some(item => !!child(child(item, 'leaderLines'), 'spPr'))) {
        warnings.add('Custom chart leader-line styles use the default browser line appearance.')
    }
    const labelPositionSettings = allDataLabelSettings.flatMap(item => [item, ...children(item, 'dLbl')])
    if (
        labelPositionSettings.some(item => {
            const position = child(item, 'dLblPos')
            return !!position && !chartDataLabelPositions.includes(position.getAttribute('val') as PptxChartDataLabelPosition)
        })
    )
        warnings.add('Unsupported chart data label positions use the default placement.')
    if (allDataLabelSettings.some(item => !!child(item, 'spPr'))) warnings.add('Chart data label backgrounds, outlines, and shape effects are not rendered.')
    if (seriesNodes.length > 12) warnings.add('Chart series beyond 12 are omitted.')
    if (isBubble && seriesNodes.some(item => chartBoolean(child(item, 'bubble3D')) === true)) warnings.add('Three-dimensional bubbles are approximated as flat circles.')
    if (isBubble) warnings.add('Bubble radii preserve relative area or width and c:bubbleScale, with the Office default pixel size approximated.')
    const series = seriesNodes
        .slice(0, 12)
        .map((item, index): PptxChartSeries => {
            const descriptor = seriesDescriptors[index]!
            const sourceIndex = Number(child(item, 'idx')?.getAttribute('val'))
            const legendIndex = Number.isSafeInteger(sourceIndex) && sourceIndex >= 0 ? sourceIndex : index
            if (child(item, 'idx') && legendIndex !== sourceIndex) warnings.add('A chart series with an invalid index uses its source order.')
            const values = parseNumericValues(child(item, isXYChart ? 'yVal' : 'val'))
            const xValues = isXYChart ? parseNumericValues(child(item, 'xVal')) : undefined
            const seriesNumberSource = child(item, isXYChart ? 'yVal' : 'val')
            const seriesNumberFormat = chartCachedNumberFormat(seriesNumberSource, warnings) ?? chartWorkbookNumberFormat(seriesNumberSource, chartWorkbook, warnings)
            const xNumberSource = isXYChart ? child(item, 'xVal') : undefined
            const xNumberFormat = xNumberSource ? chartCachedNumberFormat(xNumberSource, warnings) ?? chartWorkbookNumberFormat(xNumberSource, chartWorkbook, warnings) : undefined
            const bubbleSizes = isBubble ? parseNumericValues(child(item, 'bubbleSize')) : undefined
            if (isBubble) {
                const pointCount = Math.max(values.length, xValues?.length || 0, bubbleSizes?.length || 0)
                if (Array.from({length: pointCount}, (_, pointIndex) => values[pointIndex] != null || xValues?.[pointIndex] != null || bubbleSizes?.[pointIndex] != null).some((hasData, pointIndex) => hasData && (values[pointIndex] == null || xValues?.[pointIndex] == null || bubbleSizes?.[pointIndex] == null))) {
                    warnings.add('Bubble chart points without matching x, y, and size values are skipped.')
                }
            }
            const marker = child(item, 'marker')
            const markerSymbol = chartMarkerSymbol(marker, warnings)
            const markerSize = chartMarkerSize(marker, warnings)
            const groupDataLabels = child(descriptor.groupNode, 'dLbls')
            const groupLabelStyle = chartDataLabelTextStyle(groupDataLabels, theme, warnings)
            const groupLabelStyleOverrides = chartDataLabelTextStyleOverrides(groupDataLabels, theme, warnings)
            const groupShowMarkers = chartBoolean(child(descriptor.groupNode, 'marker')) ?? true
            const radarMarkerSize = markerSize
            const markerSymbolOverrides: Record<number, PptxChartMarkerSymbol> = {}
            const markerSizeOverrides: Record<number, number> = {}
            const lineSettings = child(child(item, 'spPr'), 'ln')
            const seriesFillColor = chartFillColor(child(item, 'spPr'), theme, warnings, palette[index % palette.length]!)
            const seriesStroke = children(lineSettings).some(node => FILL_KINDS.includes(node.localName)) ? chartFillColor(lineSettings, theme, warnings, seriesFillColor) : undefined
            const seriesStrokeWidthEmu = chartStrokeWidthEmu(lineSettings, warnings)
            const seriesStrokeDash = lineDashStyle(lineSettings, warnings)
            const seriesStrokeCap = lineCapStyle(lineSettings, warnings)
            const seriesColor = (isRadar || descriptor.chartType === 'line') && lineSettings ? chartFillColor(lineSettings, theme, warnings, seriesFillColor) : seriesFillColor
            const smoothSetting = child(item, 'smooth')
            const lineSmooth = descriptor.groupNode.localName === 'lineChart'
                ? chartBoolean(smoothSetting) ?? chartBoolean(child(descriptor.groupNode, 'smooth')) ?? false
                : undefined
            const pointColors: Record<number, string> = {}
            const pointStrokes: Record<number, string> = {}
            const pointStrokeWidthsEmu: Record<number, number> = {}
            const pointStrokeDashes: Record<number, NonNullable<PptxChartSeries['strokeDash']>> = {}
            const pointStrokeCaps: Record<number, NonNullable<PptxChartSeries['strokeCap']>> = {}
            for (const point of children(item, 'dPt')) {
                const pointIndex = numberAttr(child(point, 'idx'), 'val', -1)
                if (pointIndex < 0 || pointIndex >= 128) continue
                const pointProperties = child(point, 'spPr')
                if (pointProperties) {
                    pointColors[pointIndex] = chartFillColor(pointProperties, theme, warnings, palette[pointIndex % palette.length]!)
                    const pointLineSettings = child(pointProperties, 'ln')
                    if (children(pointLineSettings).some(node => FILL_KINDS.includes(node.localName))) pointStrokes[pointIndex] = chartFillColor(pointLineSettings, theme, warnings, pointColors[pointIndex]!)
                    const pointStrokeWidthEmu = chartStrokeWidthEmu(pointLineSettings, warnings)
                    if (pointStrokeWidthEmu !== undefined) pointStrokeWidthsEmu[pointIndex] = pointStrokeWidthEmu
                    const pointStrokeDash = lineDashStyle(pointLineSettings, warnings)
                    if (pointStrokeDash !== undefined) pointStrokeDashes[pointIndex] = pointStrokeDash
                    const pointStrokeCap = lineCapStyle(pointLineSettings, warnings)
                    if (pointStrokeCap !== undefined) pointStrokeCaps[pointIndex] = pointStrokeCap
                }
                const pointMarker = child(point, 'marker')
                const pointMarkerSymbol = chartMarkerSymbol(pointMarker, warnings)
                const pointMarkerSize = chartMarkerSize(pointMarker, warnings)
                if (pointMarkerSymbol !== undefined) markerSymbolOverrides[pointIndex] = pointMarkerSymbol
                if (pointMarkerSize !== undefined) markerSizeOverrides[pointIndex] = pointMarkerSize
            }
            const tx = child(item, 'tx')
            const name =
                descendants(tx, 't')
                    .map(value => value.textContent || '')
                    .join('') ||
                firstDescendant(tx, 'v')?.textContent ||
                `Series ${index + 1}`
            const dataLabels = child(item, 'dLbls')
            const seriesLabelStyle = chartDataLabelTextStyle(dataLabels, theme, warnings)
            const seriesLabelStyleOverrides = chartDataLabelTextStyleOverrides(dataLabels, theme, warnings)
            const groupShowLegendKey = chartDataLabelFlag(groupDataLabels, 'showLegendKey') ?? (comboMode ? false : chartShowLegendKey)
            const groupShowLeaderLines = chartDataLabelFlag(groupDataLabels, 'showLeaderLines') ?? (comboMode ? false : chartShowLeaderLines)
            const groupShowValueLabels = chartDataLabelFlag(groupDataLabels, 'showVal') ?? (comboMode ? false : chartShowValueLabels)
            const groupShowPercentLabels = chartDataLabelFlag(groupDataLabels, 'showPercent') ?? (comboMode ? false : chartShowPercentLabels)
            const groupShowBubbleSizeLabels = chartDataLabelFlag(groupDataLabels, 'showBubbleSize') ?? (comboMode ? false : chartShowBubbleSizeLabels)
            const groupShowCategoryNameLabels = chartDataLabelFlag(groupDataLabels, 'showCatName') ?? (comboMode ? false : chartShowCategoryNameLabels)
            const groupShowSeriesNameLabels = chartDataLabelFlag(groupDataLabels, 'showSerName') ?? (comboMode ? false : chartShowSeriesNameLabels)
            const groupLabelSeparator = chartDataLabelSeparator(groupDataLabels) ?? (comboMode ? undefined : chartLabelSeparator)
            const groupValueFormat = chartNumberFormat(groupDataLabels) ?? (comboMode ? undefined : chartValueFormat)
            const groupLabelPosition = chartDataLabelPosition(groupDataLabels) ?? (comboMode ? undefined : chartLabelPosition)
            const numberFormat = chartNumberFormat(dataLabels) ?? groupValueFormat ?? seriesNumberFormat
            return {
                legendIndex,
                name,
                values,
                ...(comboMode ? {
                    chartType: descriptor.chartType,
                    chartGroupIndex: descriptor.chartGroupIndex,
                    grouping: descriptor.grouping,
                    ...(descriptor.valueAxisId ? {valueAxisId: descriptor.valueAxisId} : {}),
                    ...(descriptor.chartType === 'line' ? {
                        lineShowLine: lineSettings ? !child(lineSettings, 'noFill') : true,
                        lineShowMarker: markerSymbol ? markerSymbol !== 'none' : groupShowMarkers
                    } : {})
                } : {}),
                ...(lineSmooth !== undefined ? {lineSmooth} : {}),
                ...(xValues ? {xValues} : {}),
                ...(seriesNumberFormat ? {cachedNumberFormat: seriesNumberFormat} : {}),
                ...(xNumberFormat ? {xNumberFormat} : {}),
                ...(bubbleSizes ? {bubbleSizes} : {}),
                ...(isScatter ? {
                    scatterShowLine: lineSettings ? !child(lineSettings, 'noFill') : scatterShowLine,
                    scatterShowMarker: markerSymbol ? markerSymbol !== 'none' : scatterShowMarker,
                    scatterSmooth: chartBoolean(smoothSetting) ?? scatterStyle.startsWith('smooth')
                } : {}),
                ...(isRadar ? {
                    radarShowLine: lineSettings ? !child(lineSettings, 'noFill') : true,
                    radarShowMarker: markerSymbol ? markerSymbol !== 'none' : radarStyle === 'marker',
                    ...(radarMarkerSize !== undefined && Number.isInteger(radarMarkerSize) && radarMarkerSize >= 2 && radarMarkerSize <= 72 ? {radarMarkerSize} : {})
                } : {}),
                ...(markerSymbol !== undefined ? {markerSymbol} : {}),
                ...(markerSize !== undefined ? {markerSize} : {}),
                ...(Object.keys(markerSymbolOverrides).length ? {markerSymbolOverrides} : {}),
                ...(Object.keys(markerSizeOverrides).length ? {markerSizeOverrides} : {}),
                color: seriesColor,
                ...(seriesStroke !== undefined ? {stroke: seriesStroke} : {}),
                ...(seriesStrokeWidthEmu !== undefined ? {strokeWidthEmu: seriesStrokeWidthEmu} : {}),
                ...(seriesStrokeDash !== undefined ? {strokeDash: seriesStrokeDash} : {}),
                ...(seriesStrokeCap !== undefined ? {strokeCap: seriesStrokeCap} : {}),
                pointColors,
                ...(Object.keys(pointStrokes).length ? {pointStrokes} : {}),
                ...(Object.keys(pointStrokeWidthsEmu).length ? {pointStrokeWidthsEmu} : {}),
                ...(Object.keys(pointStrokeDashes).length ? {pointStrokeDashes} : {}),
                ...(Object.keys(pointStrokeCaps).length ? {pointStrokeCaps} : {}),
                showLegendKey: chartDataLabelFlag(dataLabels, 'showLegendKey') ?? groupShowLegendKey,
                legendKeyOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showLegendKey'), ...chartDataLabelFlagOverrides(dataLabels, 'showLegendKey')},
                showValueLabels: chartDataLabelFlag(dataLabels, 'showVal') ?? groupShowValueLabels,
                valueLabelOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showVal'), ...chartDataLabelFlagOverrides(dataLabels, 'showVal')},
                dataLabelPosition: chartDataLabelPosition(dataLabels) ?? groupLabelPosition,
                dataLabelPositionOverrides: {...chartDataLabelPositionOverrides(groupDataLabels), ...chartDataLabelPositionOverrides(dataLabels)},
                showLeaderLines: chartDataLabelFlag(dataLabels, 'showLeaderLines') ?? groupShowLeaderLines,
                leaderLineOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showLeaderLines'), ...chartDataLabelFlagOverrides(dataLabels, 'showLeaderLines')},
                ...((groupLabelStyle || seriesLabelStyle) ? {dataLabelStyle: {...groupLabelStyle, ...seriesLabelStyle}} : {}),
                dataLabelStyleOverrides: mergeChartDataLabelTextStyleOverrides(groupLabelStyleOverrides, seriesLabelStyleOverrides),
                showPercentLabels: chartDataLabelFlag(dataLabels, 'showPercent') ?? groupShowPercentLabels,
                percentLabelOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showPercent'), ...chartDataLabelFlagOverrides(dataLabels, 'showPercent')},
                showBubbleSizeLabels: chartDataLabelFlag(dataLabels, 'showBubbleSize') ?? groupShowBubbleSizeLabels,
                bubbleSizeLabelOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showBubbleSize'), ...chartDataLabelFlagOverrides(dataLabels, 'showBubbleSize')},
                showCategoryNameLabels: chartDataLabelFlag(dataLabels, 'showCatName') ?? groupShowCategoryNameLabels,
                categoryNameLabelOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showCatName'), ...chartDataLabelFlagOverrides(dataLabels, 'showCatName')},
                showSeriesNameLabels: chartDataLabelFlag(dataLabels, 'showSerName') ?? groupShowSeriesNameLabels,
                seriesNameLabelOverrides: {...chartDataLabelFlagOverrides(groupDataLabels, 'showSerName'), ...chartDataLabelFlagOverrides(dataLabels, 'showSerName')},
                labelSeparator: chartDataLabelSeparator(dataLabels) ?? groupLabelSeparator,
                labelSeparatorOverrides: {...chartDataLabelSeparatorOverrides(groupDataLabels), ...chartDataLabelSeparatorOverrides(dataLabels)},
                ...(numberFormat ? {numberFormat} : {}),
                valueNumberFormatOverrides: {...chartValueNumberFormatOverrides(groupDataLabels), ...chartValueNumberFormatOverrides(dataLabels)}
            }
        })
        .filter(item => isXYChart
            ? item.values.some((value, index) => value !== null && item.xValues?.[index] != null && (!isBubble || item.bubbleSizes?.[index] != null))
            : item.values.some(value => value !== null))
    if (numberFormatNodes.some(item => !['0', 'false'].includes((item.getAttribute('sourceLinked') || '1').toLowerCase())) && series.some(item => !item.cachedNumberFormat && !item.numberFormat)) {
        warnings.add('Source-linked chart number formats use raw values when no supported cached or embedded workbook style is available.')
    }
    if (!series.length) {
        warnings.add('A chart without cached series values cannot be rendered.')
        return undefined
    }
    let stock: PptxChart['stock']
    if (isStock) {
        const legacyVolumeIndex = !stockVolumeMode && /volume|vol\b/i.test(series[0]?.name || '') ? 0 : undefined
        const volumeIndex = stockVolumeMode ? series.findIndex(item => item.chartType === 'bar') : legacyVolumeIndex
        const priceIndexes = stockVolumeMode
            ? series.flatMap((item, index) => item.chartType === 'stock' ? [index] : [])
            : Array.from({length: series.length - (legacyVolumeIndex === undefined ? 0 : 1)}, (_, index) => index + (legacyVolumeIndex === undefined ? 0 : 1))
        const priceSeriesCount = priceIndexes.length
        if (priceSeriesCount !== 3 && priceSeriesCount !== 4) {
            warnings.add('Stock charts require cached high-low-close or open-high-low-close series; this chart is not rendered.')
            return undefined
        }
        if (stockVolumeMode && (volumeIndex === undefined || volumeIndex < 0 || series.filter(item => item.chartType === 'bar').length !== 1)) {
            warnings.add('Stock volume combinations require one cached column series and three or four cached stock price series; this chart is not rendered.')
            return undefined
        }
        if (legacyVolumeIndex !== undefined) warnings.add('Stock volume data is not rendered; cached price series remain available.')
        const hasOpen = priceSeriesCount === 4
        const highIndex = priceIndexes[hasOpen ? 1 : 0]!
        const lowIndex = priceIndexes[hasOpen ? 2 : 1]!
        const closeIndex = priceIndexes[hasOpen ? 3 : 2]!
        const highLowLines = child(typeNode, 'hiLowLines')
        const highLowLine = child(child(highLowLines, 'spPr'), 'ln')
        const highLowSolidFill = child(highLowLine, 'solidFill')
        const upDownBars = child(typeNode, 'upDownBars')
        const gapWidthNode = child(upDownBars, 'gapWidth')
        const rawGapWidth = gapWidthNode ? numberAttr(gapWidthNode, 'val', 150) : 150
        const gapWidth = Number.isInteger(rawGapWidth) && rawGapWidth >= 0 && rawGapWidth <= 500 ? rawGapWidth : 150
        if (gapWidth !== rawGapWidth) warnings.add('Invalid stock up/down bar gap width uses the 150 percent default.')
        const volumeGapWidthNode = stockVolumeMode ? child(volumeChartNode, 'gapWidth') : undefined
        const rawVolumeGapWidth = volumeGapWidthNode ? numberAttr(volumeGapWidthNode, 'val', 150) : 150
        const volumeGapWidth = Number.isInteger(rawVolumeGapWidth) && rawVolumeGapWidth >= 0 && rawVolumeGapWidth <= 500 ? rawVolumeGapWidth : 150
        if (volumeGapWidth !== rawVolumeGapWidth) warnings.add('Invalid stock volume gap width uses the 150 percent default.')
        const closeColor = series[closeIndex]?.color || series[highIndex]?.color || palette[0]!
        stock = {
            highIndex,
            lowIndex,
            closeIndex,
            ...(hasOpen ? {openIndex: priceIndexes[0]} : {}),
            ...(volumeIndex !== undefined && volumeIndex >= 0 ? {volumeIndex} : {}),
            ...(stockVolumeMode ? {volumeGapWidth} : {}),
            highLowColor: (highLowSolidFill && readColor(children(highLowSolidFill)[0], theme)) || series[highIndex]?.color || closeColor,
            showHighLowLines: !child(highLowLine, 'noFill'),
            showDropLines: !!child(typeNode, 'dropLines'),
            showUpDownBars: !!upDownBars && hasOpen,
            upColor: chartFillColor(child(child(upDownBars, 'upBars'), 'spPr'), theme, warnings, '#dcfce7'),
            downColor: chartFillColor(child(child(upDownBars, 'downBars'), 'spPr'), theme, warnings, '#fecaca'),
            gapWidth,
        }
        if (upDownBars && !hasOpen) warnings.add('Stock up/down bars require open and close price series; they were omitted.')
    }
    if (isRadar && radarStyle !== 'filled' && series.length === 1 && chartBoolean(child(typeNode, 'varyColors')) === true) warnings.add('Radar point colors that vary within a single series are approximated with one series color.')
    const stacking = grouping === 'clustered' ? undefined : getPptxChartStacking(series, grouping)
    if (type === 'pie' && series.length > 1) warnings.add('Pie chart series beyond the first are omitted.')
    const valueAxisNodes = children(plotArea, 'valAx')
    let scatterXAxisNode: Element | undefined
    let scatterYAxisNode: Element | undefined
    if (isXYChart) {
        const referencedAxes = children(typeNode, 'axId')
            .map(node => valueAxisNodes.find(axis => child(axis, 'axId')?.getAttribute('val') === node.getAttribute('val')))
            .filter((axis): axis is Element => !!axis)
        const axisPosition = (axis: Element) => child(axis, 'axPos')?.getAttribute('val')
        scatterXAxisNode = valueAxisNodes.find(axis => ['b', 't'].includes(axisPosition(axis) || '')) || referencedAxes[0] || valueAxisNodes[0]
        scatterYAxisNode = valueAxisNodes.find(axis => axis !== scatterXAxisNode && ['l', 'r'].includes(axisPosition(axis) || ''))
            || referencedAxes.find(axis => axis !== scatterXAxisNode)
            || valueAxisNodes.find(axis => axis !== scatterXAxisNode)
        if (valueAxisNodes.length < 2 || !scatterXAxisNode || !scatterYAxisNode) warnings.add('An XY chart is missing one of its numeric axes; available axes are used.')
        if (valueAxisNodes.length > 2) warnings.add('Secondary XY chart axes are omitted.')
        const xId = child(scatterXAxisNode, 'axId')?.getAttribute('val')
        const yId = child(scatterYAxisNode, 'axId')?.getAttribute('val')
        const xCross = child(scatterXAxisNode, 'crossAx')?.getAttribute('val')
        const yCross = child(scatterYAxisNode, 'crossAx')?.getAttribute('val')
        if ((xCross && yId && xCross !== yId) || (yCross && xId && yCross !== xId)) warnings.add('XY chart axis cross-references are invalid; axis positions are used.')
    }
    const sourceNumberFormatForAxis = (axisId?: string) => sharedChartNumberFormat(series
        .filter(item => !axisId || !item.valueAxisId || item.valueAxisId === axisId)
        .map(item => item.cachedNumberFormat))
    const parsedXAxis = isXYChart ? chartValueAxis(scatterXAxisNode, warnings, 'x', sharedChartNumberFormat(series.map(item => item.xNumberFormat))) : undefined
    const primaryValueAxisNode = valueAxisNodes.find(axis => child(axis, 'axId')?.getAttribute('val') === primaryValueAxisId) || valueAxisNodes[0]
    const parsedValueAxis = isPieLike ? undefined : chartValueAxis(isXYChart ? scatterYAxisNode : primaryValueAxisNode, warnings, 'value', sourceNumberFormatForAxis(primaryValueAxisId))
    const categoryAxisNodes = isPieLike || isXYChart ? [] : children(plotArea, 'catAx')
    const categoryAxis = categoryAxisNodes.find(axis => child(axis, 'axId')?.getAttribute('val') === primaryGroupAxisIds[0]) || categoryAxisNodes[0]
    const dateAxisNodes = isPieLike || isXYChart ? [] : children(plotArea, 'dateAx')
    const dateAxis = dateAxisNodes[0]
    const dateAxisScaling = child(dateAxis, 'scaling')
    const dateAxisBound = (name: 'min' | 'max') => {
        const raw = child(dateAxisScaling, name)?.getAttribute('val')
        const value = raw === null || raw === undefined ? NaN : Number(raw)
        return Number.isFinite(value) ? value : undefined
    }
    const dateAxisMinimum = dateAxisBound('min')
    const dateAxisMaximum = dateAxisBound('max')
    const dateAxisReverse = child(dateAxisScaling, 'orientation')?.getAttribute('val') === 'maxMin'
    const dateAxisMajorUnitNode = child(dateAxis, 'majorUnit')
    const dateAxisMajorUnitRaw = dateAxisMajorUnitNode?.getAttribute('val')
    const dateAxisMajorUnitValue = dateAxisMajorUnitRaw === null || dateAxisMajorUnitRaw === undefined ? NaN : Number(dateAxisMajorUnitRaw)
    const dateAxisMajorUnit = Number.isFinite(dateAxisMajorUnitValue) && dateAxisMajorUnitValue > 0 ? dateAxisMajorUnitValue : undefined
    const dateAxisMajorTimeUnitNode = child(dateAxis, 'majorTimeUnit')
    const dateAxisMajorTimeUnitValue = dateAxisMajorTimeUnitNode?.getAttribute('val')
    const dateAxisMajorTimeUnit = dateAxisMajorTimeUnitValue === 'days' || dateAxisMajorTimeUnitValue === 'months' || dateAxisMajorTimeUnitValue === 'years'
        ? dateAxisMajorTimeUnitValue
        : undefined
    const dateAxisMinorUnitNode = child(dateAxis, 'minorUnit')
    const dateAxisMinorUnitValue = dateAxisMinorUnitNode?.getAttribute('val')
    const parsedDateAxisMinorUnit = dateAxisMinorUnitValue === null || dateAxisMinorUnitValue === undefined ? NaN : Number(dateAxisMinorUnitValue)
    const dateAxisMinorUnit = Number.isFinite(parsedDateAxisMinorUnit) && parsedDateAxisMinorUnit > 0 ? parsedDateAxisMinorUnit : undefined
    const dateAxisMinorTimeUnitValue = child(dateAxis, 'minorTimeUnit')?.getAttribute('val')
    const dateAxisMinorTimeUnit = dateAxisMinorTimeUnitValue === 'days' || dateAxisMinorTimeUnitValue === 'months' || dateAxisMinorTimeUnitValue === 'years'
        ? dateAxisMinorTimeUnitValue
        : undefined
    const categoryScaleAxis = categoryAxis || dateAxis
    const categoryAxisNumberFormatNode = child(categoryAxis || dateAxis, 'numFmt')
    const categoryAxisFormatCode = categoryAxisNumberFormatNode?.getAttribute('formatCode')?.trim()
    const categoryAxisSourceLinked = (categoryAxisNumberFormatNode?.getAttribute('sourceLinked') || '1').toLowerCase()
    const categorySource = child(seriesNodes[0], 'cat')
    const categorySourceNumberFormat = chartCachedNumberFormat(categorySource, warnings) ?? chartWorkbookNumberFormat(categorySource, chartWorkbook, warnings)
    let categoryAxisNumberFormat: string | undefined
    if (categoryAxisFormatCode && ['0', 'false'].includes(categoryAxisSourceLinked)) {
        if (isSupportedChartNumberFormat(categoryAxisFormatCode)) categoryAxisNumberFormat = categoryAxisFormatCode
        else warnings.add('Unsupported chart number formats are rendered as raw values.')
    } else if (categoryAxisFormatCode) {
        if (categorySourceNumberFormat) categoryAxisNumberFormat = categorySourceNumberFormat
        else warnings.add('Source-linked chart number formats use raw values when no supported cached or embedded workbook style is available.')
    }
    if (dateAxis && (dateAxisMajorUnitNode || dateAxisMajorTimeUnitNode)) {
        if (!dateAxisMajorUnit) warnings.add('A date-axis c:majorUnit must be a positive number; cached category labels are used.')
        if (!dateAxisMajorTimeUnit) warnings.add('A date-axis c:majorTimeUnit must be days, months, or years; cached category labels are used.')
        if (!dateAxisMajorUnit || !dateAxisMajorTimeUnit) warnings.add('Explicit date-axis major ticks require both a valid c:majorUnit and c:majorTimeUnit.')
    } else if (dateAxis) {
        warnings.add('Automatic date-axis major intervals are approximated; cached category labels are used.')
    }
    if (dateAxis && (dateAxisMinorUnitNode || child(dateAxis, 'minorTimeUnit')) && (!dateAxisMinorUnit || !dateAxisMinorTimeUnit || (dateAxisMinorTimeUnit !== 'days' && !Number.isSafeInteger(dateAxisMinorUnit)))) {
        warnings.add('Date-axis minor ticks require a supported positive c:minorUnit and c:minorTimeUnit; automatic subdivisions are used.')
    }
    if (dateAxisNodes.length > 1) warnings.add('Secondary chart date axes are omitted.')
    let categoryAxisPosition = isPieLike || isXYChart ? undefined : chartAxisPosition(categoryScaleAxis, 'category', warnings)
    const categoryAxisMajorTickMark = isPieLike || isXYChart ? undefined : chartAxisTickMark(categoryScaleAxis, 'category', 'major', warnings)
    const categoryAxisMinorTickMark = isPieLike || isXYChart ? undefined : chartAxisTickMark(categoryScaleAxis, 'category', 'minor', warnings)
    const categoryAxisTickLabelSkip = categoryAxis ? chartAxisSkip(categoryAxis, 'tickLblSkip', warnings) : undefined
    const categoryAxisTickMarkSkip = categoryAxis ? chartAxisSkip(categoryAxis, 'tickMarkSkip', warnings) : undefined
    if (dateAxis && categoryAxisMinorTickMark && categoryAxisMinorTickMark !== 'none' && (!dateAxisMinorUnit || !dateAxisMinorTimeUnit)) {
        warnings.add('Automatic date-axis minor intervals use four subdivisions per major interval.')
    }
    const categoryAxisTickLabelPosition = isPieLike || isXYChart ? undefined : chartTickLabelPosition(categoryScaleAxis, 'category', warnings)
    const categoryAxisDeleted = !isPieLike && !isXYChart && chartAxisDeleted(categoryScaleAxis, 'category', warnings)
    let xAxisPosition = isXYChart ? chartAxisPosition(scatterXAxisNode, 'x', warnings) : undefined
    if (isXYChart && xAxisPosition && !['b', 't'].includes(xAxisPosition)) {
        warnings.add('An XY chart x-axis must be positioned on the top or bottom; its default side is used.')
        xAxisPosition = undefined
    }
    if (direction === 'horizontal' && !isScatter && parsedValueAxis?.position && !['b', 't'].includes(parsedValueAxis.position)) {
        warnings.add('A horizontal chart value axis must be positioned on the top or bottom; its default side is used.')
        delete parsedValueAxis.position
    }
    if ((direction === 'vertical' || isXYChart) && parsedValueAxis?.position && !['l', 'r'].includes(parsedValueAxis.position)) {
        warnings.add('A vertical chart value axis must be positioned on the left or right; its default side is used.')
        delete parsedValueAxis.position
    }
    if (categoryAxisPosition && (direction === 'horizontal' ? !['l', 'r'].includes(categoryAxisPosition) : !['b', 't'].includes(categoryAxisPosition))) {
        warnings.add('The chart category axis position does not match its chart direction; its default side is used.')
        categoryAxisPosition = undefined
    }
    if (!comboMode && !isPieLike && !isXYChart && valueAxisNodes.length > 1) warnings.add('Secondary chart value axes are omitted.')
    if (!isPieLike && !isXYChart && categoryAxisNodes.length > 1) warnings.add('Secondary chart category axes are omitted.')
    const xyPoints = isXYChart ? series.flatMap(item => item.values.flatMap((value, index) => {
        const x = item.xValues?.[index]
        return value !== null && x !== null && x !== undefined ? [{x, y: value}] : []
    })) : []
    const resolveAxis = (axis: ReturnType<typeof chartValueAxis>, values: number[], axisName: 'value' | 'x', includeZero: boolean): PptxChartAxis | undefined => {
        if (!axis) return undefined
        let logBase = axis.logBase
        let rangeValues = values
        if (logBase !== undefined) {
            const positiveValues = values.filter(value => value > 0)
            if (axis.min !== undefined && axis.min <= 0) {
                warnings.add(`A logarithmic chart ${axisName}-axis requires a positive minimum; the automatic minimum is used.`)
                delete axis.min
            }
            if (axis.max !== undefined && axis.max <= 0) {
                warnings.add(`A logarithmic chart ${axisName}-axis requires a positive maximum; the automatic maximum is used.`)
                delete axis.max
            }
            if (!positiveValues.length && (axis.min === undefined || axis.max === undefined)) {
                warnings.add(`A logarithmic chart ${axisName}-axis without positive cached values and two explicit positive bounds is rendered linearly.`)
                delete axis.logBase
                logBase = undefined
            } else {
                rangeValues = positiveValues
            }
        }
        const dataMin = rangeValues.length ? Math.min(...rangeValues) : 0
        const dataMax = rangeValues.length ? Math.max(...rangeValues) : 0
        const roundToLogPower = (value: number, direction: 'floor' | 'ceil') => {
            if (!(value > 0)) return undefined
            const exponent = Math.log(value) / Math.log(logBase!)
            const rounded = Math.pow(logBase!, direction === 'floor' ? Math.floor(exponent + 1e-12) : Math.ceil(exponent - 1e-12))
            return Number.isFinite(rounded) && rounded > 0 ? rounded : value
        }
        const automaticLogMinimum = logBase !== undefined && dataMin > 0 ? roundToLogPower(dataMin, 'floor')! : 1
        const automaticLogMaximum = logBase !== undefined && dataMax > 0 ? roundToLogPower(dataMax, 'ceil')! : automaticLogMinimum * (logBase ?? 10)
        let min = axis.min ?? (logBase !== undefined ? automaticLogMinimum : includeZero ? Math.min(0, dataMin) : dataMin)
        let max = axis.max ?? (logBase !== undefined ? automaticLogMaximum : includeZero ? Math.max(0, dataMax) : dataMax)
        if (min >= max) {
            warnings.add(`Chart ${axisName}-axis bounds conflict with the data range; data-driven bounds are used.`)
            min = logBase !== undefined ? automaticLogMinimum : includeZero ? Math.min(0, dataMin) : dataMin
            max = logBase !== undefined ? automaticLogMaximum : includeZero ? Math.max(0, dataMax) : dataMax
            delete axis.min
            delete axis.max
        }
        if (min === max) {
            if (logBase !== undefined) {
                const expanded = min * logBase
                if (Number.isFinite(expanded) && expanded > min) max = expanded
                else min /= logBase
            } else if (includeZero) max = min + 1
            else {
                const padding = Math.abs(min) * 0.05 || 0.5
                min -= padding
                max += padding
            }
        }
        if (axis.crossesAt !== undefined && (axis.crossesAt < min || axis.crossesAt > max)) warnings.add(`A chart ${axisName}-axis crossing value is outside the displayed range and is clamped to the nearest bound.`)
        const intervalRange = logBase !== undefined ? Math.log(max / min) / Math.log(logBase) : max - min
        if (axis.majorUnit && Math.floor(intervalRange / axis.majorUnit + 1e-9) > 100) {
            warnings.add(`Chart ${axisName}-axis major units produce too many ticks and are ignored.`)
            delete axis.majorUnit
        }
        if (axis.minorUnit && axis.minorTickMark !== 'none' && Math.floor(intervalRange / axis.minorUnit + 1e-9) > 512) {
            warnings.add(`Chart ${axisName}-axis minor units produce too many ticks and are ignored.`)
            delete axis.minorUnit
        }
        return {...axis, min, max}
    }
    let valueAxis: PptxChart['valueAxis']
    let valueAxes: PptxChart['valueAxes']
    const axisValues = isXYChart ? xyPoints.map(point => point.y) : (stock ? [stock.highIndex, stock.lowIndex, stock.closeIndex, stock.openIndex].filter((index): index is number => index !== undefined).flatMap(index => series[index]?.values.filter((value): value is number => value !== null) || []) : series.flatMap(item => item.values.filter((value): value is number => value !== null)))
    if (comboMode) {
        if (valueAxisNodes.length > 2) warnings.add('Combo charts with more than two value axes use the first two referenced axes.')
        const primaryId = primaryValueAxisId || child(primaryValueAxisNode, 'axId')?.getAttribute('val') || ''
        const axisIds = new Set(valueAxisNodes.map(axis => child(axis, 'axId')?.getAttribute('val')).filter((id): id is string => !!id))
        for (const item of series) {
            if (!item.valueAxisId || !axisIds.has(item.valueAxisId)) {
                if (item.valueAxisId) warnings.add('A combo chart series references a missing value axis; the primary value axis is used.')
                item.valueAxisId = primaryId
            }
        }
        const selectedValueAxisNodes = [
            ...valueAxisNodes.filter(axis => child(axis, 'axId')?.getAttribute('val') === primaryId),
            ...valueAxisNodes.filter(axis => child(axis, 'axId')?.getAttribute('val') !== primaryId)
        ].slice(0, 2)
        valueAxes = selectedValueAxisNodes.flatMap(axis => {
            const id = child(axis, 'axId')?.getAttribute('val')
            if (!id) return []
            const axisSeries = series.filter(item => item.valueAxisId === id)
            const seriesGroups = new Map<number, PptxChartSeries[]>()
            for (const item of axisSeries) {
                const group = item.chartGroupIndex ?? 0
                seriesGroups.set(group, [...(seriesGroups.get(group) || []), item])
            }
            const axisValues = [...seriesGroups.values()].flatMap(group => {
                const grouping = group[0]?.grouping
                if (!grouping || grouping === 'clustered') return group.flatMap(item => item.values.filter((value): value is number => value !== null))
                const {segments} = getPptxChartStacking(group, grouping)
                return segments.flatMap(items => items.flatMap(segment => segment ? [segment.start, segment.end] : []))
            })
            const parsedAxis = id === primaryId ? parsedValueAxis : chartValueAxis(axis, warnings, 'value', sourceNumberFormatForAxis(id))
            const resolved = resolveAxis(parsedAxis, axisValues, 'value', true)
            return resolved ? [{id, axis: resolved, ...(id === primaryId ? {primary: true} : {})}] : []
        })
        valueAxis = valueAxes.find(item => item.primary)?.axis || resolveAxis(parsedValueAxis, axisValues, 'value', true)
    } else {
        valueAxis = resolveAxis(parsedValueAxis, axisValues, 'value', !isXYChart)
    }
    const xAxis = resolveAxis(parsedXAxis, xyPoints.map(point => point.x), 'x', false)
    let omittedLogAxisValues = false
    for (const item of series) {
        const itemValueAxis = valueAxes?.find(axis => axis.id === item.valueAxisId)?.axis || valueAxis
        if (itemValueAxis?.logBase) {
            item.values = item.values.map(value => {
                if (value !== null && value <= 0) {
                    omittedLogAxisValues = true
                    return null
                }
                return value
            })
        }
        if (xAxis?.logBase && item.xValues) {
            item.xValues = item.xValues.map(value => {
                if (value !== null && value <= 0) {
                    omittedLogAxisValues = true
                    return null
                }
                return value
            })
        }
    }
    if (omittedLogAxisValues) warnings.add('Non-positive cached values on logarithmic chart axes are omitted because log scales require positive values.')
    const categories = isXYChart ? [] : chartDataValues(child(seriesNodes[0], 'cat'), chartWorkbook, warnings)
    const pointCount = isXYChart ? 0 : Math.min(128, Math.max(categories.length, ...series.map(item => item.values.length), 0))
    const labels = Array.from({length: pointCount}, (_, index) => categories[index] || String(index + 1))
    const dateAxisValues = dateAxis && categories.length > 0 && categories.length === pointCount
        && categories.every(category => category !== null && category.trim() !== '' && Number.isFinite(Number(category)))
        ? categories.map(category => Number(category))
        : undefined
    if (dateAxis && !dateAxisValues) warnings.add('Date-axis categories without complete numeric cached dates use evenly spaced positions.')
    if (dateAxis && dateAxisValues && dateAxisMajorUnit !== undefined && dateAxisMajorTimeUnit) {
        const minimum = dateAxisMinimum ?? Math.min(...dateAxisValues)
        const maximum = dateAxisMaximum ?? Math.max(...dateAxisValues)
        if (!(maximum > minimum) || !getChartDateAxisMajorTicks(minimum, maximum, dateAxisMajorUnit, dateAxisMajorTimeUnit, chartDate1904)) {
            warnings.add('Date-axis major interval is outside the supported calendar range or exceeds 101 ticks; cached category labels are used.')
        }
    }
    if (dateAxis && dateAxisValues && dateAxisMinorUnit !== undefined && dateAxisMinorTimeUnit) {
        const minimum = dateAxisMinimum ?? Math.min(...dateAxisValues)
        const maximum = dateAxisMaximum ?? Math.max(...dateAxisValues)
        if (!(maximum > minimum) || !getChartDateAxisMajorTicks(minimum, maximum, dateAxisMinorUnit, dateAxisMinorTimeUnit, chartDate1904)) {
            warnings.add('Date-axis minor interval is outside the supported calendar range or exceeds 101 ticks; automatic subdivisions are used.')
        }
    }
    if (isRadar && series.some(item => labels.some((_label, index) => item.values[index] == null))) warnings.add('Radar chart gaps are left open; filled radar areas require every category value.')
    if (isRadar && labels.length < 3) warnings.add('Radar charts with fewer than three categories have limited geometry.')
    if (stock && labels.some((_label, index) => [stock.highIndex, stock.lowIndex, stock.closeIndex, stock.openIndex].some(seriesIndex => seriesIndex !== undefined && series[seriesIndex]?.values[index] == null))) warnings.add('Stock chart categories with missing price values are omitted.')
    const categoryAxisCrossing: ReturnType<typeof chartAxisCrossing> = isPieLike ? {} : chartAxisCrossing(categoryAxis, 'category', warnings, Math.max(labels.length, 1))
    if (!isXYChart && labels.length !== categories.length) warnings.add('Missing chart category labels are shown as point numbers.')
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
            ...(['bar', 'line', 'area'].includes(type) ? {grouping} : {}),
            ...(isRadar ? {radarStyle} : {}),
            ...(stock ? {stock} : {}),
            ...(type === 'doughnut' ? {holeSize} : {}),
            ...(firstSliceAngleNode ? {firstSliceAngle} : {}),
            ...(isBubble ? {bubbleScale, bubbleSizeRepresents, showNegativeBubbles, bubbleVaryColors} : {}),
            title,
            categories: labels,
            palette,
            ...(indexedColors ? {indexedColors} : {}),
            series,
            ...(chartLabelSeparator !== undefined ? {labelSeparator: chartLabelSeparator} : {}),
            ...(chartValueFormat ? {numberFormat: chartValueFormat} : {}),
            ...(chartDate1904 ? {date1904: true} : {}),
            ...(dateAxisValues ? {dateAxis: {
                values: dateAxisValues,
                ...(dateAxisMinimum !== undefined ? {min: dateAxisMinimum} : {}),
                ...(dateAxisMaximum !== undefined ? {max: dateAxisMaximum} : {}),
                ...(dateAxisReverse ? {reverse: true} : {}),
                ...(dateAxisMajorUnit !== undefined ? {majorUnit: dateAxisMajorUnit} : {}),
                ...(dateAxisMajorTimeUnit ? {majorTimeUnit: dateAxisMajorTimeUnit} : {}),
                ...(dateAxisMinorUnit && dateAxisMinorTimeUnit && getChartDateAxisMajorTicks(dateAxisMinimum ?? Math.min(...dateAxisValues), dateAxisMaximum ?? Math.max(...dateAxisValues), dateAxisMinorUnit, dateAxisMinorTimeUnit, chartDate1904) ? {minorUnit: dateAxisMinorUnit, minorTimeUnit: dateAxisMinorTimeUnit} : {}),
            }} : {}),
            ...(valueAxis ? {valueAxis} : {}),
            ...(xAxis ? {xAxis: {...xAxis, ...(xAxisPosition ? {position: xAxisPosition} : {})}} : {}),
            ...(valueAxes ? {valueAxes} : {}),
            ...(categoryAxisNumberFormat ? {categoryAxisNumberFormat} : {}),
            ...(categoryAxisPosition ? {categoryAxisPosition} : {}),
            ...(categoryAxisDeleted ? {categoryAxisDeleted: true} : {}),
            ...(categoryScaleAxis && child(categoryScaleAxis, 'majorGridlines') ? {categoryAxisMajorGridlines: true} : {}),
            ...(categoryAxisMajorTickMark ? {categoryAxisMajorTickMark} : {}),
            ...(categoryAxisMinorTickMark ? {categoryAxisMinorTickMark} : {}),
            ...(categoryAxisTickLabelPosition ? {categoryAxisTickLabelPosition} : {}),
            ...(categoryAxisTickLabelSkip ? {categoryAxisTickLabelSkip} : {}),
            ...(categoryAxisTickMarkSkip ? {categoryAxisTickMarkSkip} : {}),
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

type ParsedTheme = {
    colors: Record<string, string>
    fillStyles: Element[]
    lineStyles: Element[]
    effectStyles: Element[]
    backgroundFillStyles: Element[]
}

const DEFAULT_COLOR_MAP: Record<string, string> = {
    bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
    accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
    accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
    hlink: 'hlink', folHlink: 'folHlink'
}

function resolveColorMap(root: Element | undefined, inherited: Record<string, string>, master: Record<string, string>, warnings: Set<string>): Record<string, string> {
    // An explicit masterClrMapping selects the master, even when the layout has an override.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.colormapoverride
    const override = child(root, 'clrMapOvr')
    const mapping = child(root, 'clrMap') || child(override, 'overrideClrMapping')
    if (!mapping) {
        if (child(override, 'masterClrMapping')) return master
        if (override) warnings.add('An empty or unsupported color-map override inherited its parent mapping.')
        return inherited
    }
    return readColorMap(mapping, inherited, warnings)
}

function readColorMap(mapping: Element, inherited: Record<string, string>, warnings: Set<string>): Record<string, string> {
    const result = {...inherited}
    for (const slot of Object.keys(DEFAULT_COLOR_MAP)) {
        const target = mapping.getAttribute(slot)
        if (target === null) warnings.add('An incomplete color map inherited its unspecified entries.')
        else if (!Object.hasOwn(DEFAULT_THEME, target)) warnings.add(`The invalid color-map target "${target}" for ${slot} inherited its parent entry.`)
        else result[slot] = target
    }
    return result
}

function mappedThemeColors(colors: Record<string, string>, mapping: Record<string, string>, inherited: Record<string, string> = colors): Record<string, string> {
    const result = {...inherited}
    // Read every target from the original theme so accent swaps are simultaneous, not recursive.
    for (const [slot, target] of Object.entries(mapping)) result[slot] = colors[target] || DEFAULT_THEME[target] || '#000000'
    return result
}

function parseTheme(xml: XMLDocument | undefined): ParsedTheme {
    const result = {...DEFAULT_THEME}
    const scheme = xml && firstDescendant(xml.documentElement, 'clrScheme')
    for (const slot of children(scheme)) {
        const color = children(slot)[0]
        if (!color) continue
        const value = readColor(color, result)
        if (value) result[slot.localName] = value
    }
    const fonts = xml && firstDescendant(xml.documentElement, 'fontScheme')
    for (const role of ['major', 'minor']) {
        const fontSet = child(fonts, `${role}Font`)
        for (const script of ['latin', 'ea', 'cs']) {
            const typeface = child(fontSet, script)?.getAttribute('typeface')?.trim()
            if (typeface) result[`font-${role}-${script}`] = typeface
        }
        for (const font of children(fontSet, 'font')) {
            const script = font.getAttribute('script')?.trim()
            const typeface = font.getAttribute('typeface')?.trim()
            if (script && /^[A-Za-z]{4}$/.test(script) && typeface) result[`font-${role}-script-${script[0]!.toUpperCase()}${script.slice(1).toLowerCase()}`] = typeface
        }
    }
    const formatScheme = firstDescendant(xml?.documentElement, 'fmtScheme')
    return {
        colors: result,
        fillStyles: children(child(formatScheme, 'fillStyleLst')),
        lineStyles: children(child(formatScheme, 'lnStyleLst')),
        effectStyles: children(child(formatScheme, 'effectStyleLst')),
        backgroundFillStyles: children(child(formatScheme, 'bgFillStyleLst'))
    }
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

async function resolvePictureBullets(elements: Pick<PptxElement, 'paragraphs' | 'table'>[], zip: JSZip, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>): Promise<void> {
    for (const element of elements) {
        const paragraphs = [...element.paragraphs, ...(element.table?.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs)) || [])]
        for (const paragraph of paragraphs) {
            const bullet = paragraph.bullet
            if (bullet?.kind !== 'image' || !bullet.target) continue
            bullet.url = await imageUrlFor(bullet.target, zip, urls, imageUrlCache)
            delete bullet.target
            if (!bullet.url) warnings.add('A picture bullet image could not be resolved; its marker was omitted.')
        }
    }
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

async function transitionAudioUrlFor(target: string, zip: JSZip, urls: string[], cache: Map<string, string>): Promise<string | undefined> {
    const cached = cache.get(target)
    if (cached) return cached
    const media = zip.file(target)
    const mime = mediaMime(target)
    if (!media || !mime?.startsWith('audio/')) return undefined
    try {
        const blob = await media.async('blob')
        const url = URL.createObjectURL(new Blob([blob], {type: mime}))
        cache.set(target, url)
        urls.push(url)
        return url
    } catch {
        return undefined
    }
}

async function parseTransitionSoundActions(transition: Element, rels: Map<string, PptxRelationship>, zip: JSZip, warnings: Set<string>, urls: string[], assetUrlCache: Map<string, string>): Promise<NonNullable<NonNullable<PptxSlide['transition']>['soundActions']>> {
    const actions: NonNullable<NonNullable<PptxSlide['transition']>['soundActions']> = []
    for (const soundAction of children(transition, 'sndAc').flatMap(action => children(action))) {
        if (soundAction.localName === 'endSnd') {
            actions.push({type: 'end'})
            continue
        }
        if (soundAction.localName !== 'stSnd') continue
        const sound = child(soundAction, 'snd')
        const relationshipId = namespacedAttr(sound, 'embed')
        if (!relationshipId) {
            const linkId = namespacedAttr(sound, 'link')
            warnings.add(linkId && rels.get(linkId)?.external ? 'An external transition sound was skipped; only embedded audio is supported.' : 'A transition sound without an embedded audio relationship was skipped.')
            continue
        }
        const relation = rels.get(relationshipId)
        if (!relation || relation.external) {
            warnings.add('A transition sound relationship could not be resolved to embedded audio.')
            continue
        }
        const mime = mediaMime(relation.target)
        if (!mime?.startsWith('audio/')) {
            warnings.add('A transition sound has an unsupported or invalid audio file type.')
            continue
        }
        const url = await transitionAudioUrlFor(relation.target, zip, urls, assetUrlCache)
        if (!url) {
            warnings.add('An embedded transition sound file could not be found.')
            continue
        }
        const rawLoop = soundAction.getAttribute('loop')?.trim().toLowerCase()
        if (rawLoop !== undefined && !['0', '1', 'false', 'true'].includes(rawLoop)) warnings.add('A transition sound has an invalid loop value; looping was disabled.')
        actions.push({type: 'start', url, loop: rawLoop === '1' || rawLoop === 'true'})
    }
    return actions
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

function selectedTransitionOnRoot(root: Element, warnings: Set<string>): {transition: Element; durationSource: Element} | undefined {
    let transition = child(root, 'transition')
    if (transition) return {transition, durationSource: transition}
    for (const alternate of children(root, 'AlternateContent')) {
        const choices = children(alternate, 'Choice')
        const choiceTransitions = choices.map(choice => ({choice, transition: child(choice, 'transition')})).filter((candidate): candidate is {choice: Element; transition: Element} => Boolean(candidate.transition))
        transition = child(child(alternate, 'Fallback'), 'transition')
        if (!choiceTransitions.length && !transition) continue
        const supportedChoices = choiceTransitions.filter(({choice}) => {
            const requiredPrefixes = (choice.getAttribute('Requires') || '').trim().split(/\s+/).filter(Boolean)
            if (!requiredPrefixes.length) {
                warnings.add('An AlternateContent slide transition choice with a missing or empty Requires attribute was skipped.')
                return false
            }
            return requiredPrefixes.every(prefix => SUPPORTED_MC_REQUIRES_NAMESPACES.has(choice.lookupNamespaceURI(prefix) || ''))
        })
        const selectedChoice = supportedChoices
            .map(candidate => candidate.transition)
            .find(candidate => {
                const effect = children(candidate).find(node => !['extLst', 'sndAc', 'sound'].includes(node.localName))
                return effect ? SUPPORTED_TRANSITIONS.has(effect.localName) : children(candidate, 'sndAc').length > 0
            })
        if (selectedChoice) return {transition: selectedChoice, durationSource: selectedChoice}
        const durationSource = choiceTransitions.map(candidate => candidate.transition).find(candidate => candidate.hasAttributeNS(POWERPOINT_2010_NS, 'dur')) || transition
        if (choices.length && transition) warnings.add('An unsupported AlternateContent slide transition choice was skipped in favor of its fallback.')
        if (transition) return {transition, durationSource: durationSource || transition}
    }
    return undefined
}

function selectedTransition(slideRoot: Element, warnings: Set<string>, layoutRoot?: Element, masterRoot?: Element): ({transition: Element; durationSource: Element; sourceRoot: Element}) | undefined {
    // Slide-local transition data overrides its layout, which overrides the slide master.
    for (const sourceRoot of [slideRoot, layoutRoot, masterRoot]) {
        if (!sourceRoot) continue
        const selection = selectedTransitionOnRoot(sourceRoot, warnings)
        if (selection) return {...selection, sourceRoot}
    }
    return undefined
}

function parseSlideAdvance(slideRoot: Element, warnings: Set<string>, layoutRoot?: Element, masterRoot?: Element): Pick<PptxSlide, 'advanceOnClick' | 'advanceAfterMs'> {
    const transition = selectedTransition(slideRoot, warnings, layoutRoot, masterRoot)?.transition
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

function parsePresentationShowSettings(showPr: Element | undefined, warnings: Set<string>): {loop: boolean; showAnimation: boolean; showNarration: boolean; useTimings: boolean} {
    const readBoolean = (name: 'loop' | 'showAnimation' | 'showNarration' | 'useTimings', fallback: boolean) => {
        const value = showPr?.getAttribute(name)?.trim().toLowerCase()
        if (value === undefined) return fallback
        if (value === '1' || value === 'true') return true
        if (value === '0' || value === 'false') return false
        warnings.add(`The presentation show setting ${name} is invalid; its default was used.`)
        return fallback
    }
    return {
        loop: readBoolean('loop', false),
        showAnimation: readBoolean('showAnimation', true),
        showNarration: readBoolean('showNarration', false),
        useTimings: readBoolean('useTimings', true),
    }
}

function parseSlideHidden(slideRoot: Element, warnings: Set<string>): boolean | undefined {
    const value = slideRoot.getAttribute('show')?.trim().toLowerCase()
    if (value === undefined) return undefined
    if (value === '0' || value === 'false') return true
    if (value === '1' || value === 'true') return undefined
    warnings.add('A slide has an invalid show attribute and was treated as visible.')
    return undefined
}

async function parseTransition(slideRoot: Element, warnings: Set<string>, rels: Map<string, PptxRelationship>, zip: JSZip, urls: string[], assetUrlCache: Map<string, string>, layoutRoot?: Element, layoutRels?: Map<string, PptxRelationship>, masterRoot?: Element, masterRels?: Map<string, PptxRelationship>): Promise<PptxSlide['transition']> {
    const selection = selectedTransition(slideRoot, warnings, layoutRoot, masterRoot)
    if (!selection) return undefined
    const {transition, durationSource} = selection
    const transitionRels = selection.sourceRoot === layoutRoot ? layoutRels || rels : selection.sourceRoot === masterRoot ? masterRels || rels : rels
    const soundActions = await parseTransitionSoundActions(transition, transitionRels, zip, warnings, urls, assetUrlCache)
    const effectNode = children(transition).find(node => !['extLst', 'sndAc', 'sound'].includes(node.localName))
    if (!effectNode) return soundActions.length ? {effect: 'cut', direction: 'l', durationMs: 0, soundActions} : undefined
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
    // ECMA-376 CT_SlideTransition declares spd default="fast"; invalid values fall back to that default.
    const rawSpeed = transition.hasAttribute('spd') ? trimXmlWhitespace(transition.getAttribute('spd')) : undefined
    const speed = rawSpeed === 'fast' || rawSpeed === 'med' || rawSpeed === 'slow' ? rawSpeed : 'fast'
    if (rawSpeed !== undefined && speed !== rawSpeed) warnings.add(`The slide transition speed "${rawSpeed}" is invalid; it was ignored.`)
    // OOXML defines spd as a category, not a duration in milliseconds. p14:dur is the explicit duration;
    // see Microsoft's example where p14:dur="1500" means 1.5 seconds. The values below are player fallbacks only.
    // https://ecma-international.org/publications-and-standards/standards/ecma-376/
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/99b95b35-568a-4652-9cba-df3d1175952f
    const fallbackDurationMs = {fast: 500, med: 750, slow: 1000}[speed]
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
    const rawDirection = effectNode?.getAttribute('dir')?.trim()
    const directionSpecified = Boolean(effectNode?.hasAttribute('dir'))
    const cardinalDirections = ['l', 'u', 'r', 'd']
    const stripsDirections = ['ld', 'lu', 'rd', 'ru']
    const sideDirections = ['l', 'r', 'u', 'd', 'ld', 'lu', 'rd', 'ru']
    const vortexDirections = ['l', 'r', 'u', 'd']
    const leftRightDirections = ['l', 'r']
    const rippleDirections = ['center', 'lu', 'ru', 'ld', 'rd']
    // ECMA-376 CT_SplitTransition defaults to dir=out and orient=horz.
    // https://ecma-international.org/publications-and-standards/standards/ecma-376/
    // MS-PPTX CT_GlitterTransition uses side direction and diamond/hexagon pattern.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/93a73ee5-6799-478f-8c1f-3db23726e10e
    const glitterPatterns = ['diamond', 'hexagon']
    // ECMA-376 Part 1 §A.3 uses SideDirection for push/wipe (l/u/r/d) and EightDirection for cover/pull;
    // both types default to l. Reject invalid enum values and preserve the schema default.
    // https://ecma-international.org/publications-and-standards/standards/ecma-376/
    // ECMA-376 Part 1 §A.3 CT_InOutTransition declares dir="out" as the default for p:zoom.
    // https://dev.ecma-international.org/publications-and-standards/standards/ecma-376/
    // ECMA-376 Part 1 §A.3 CT_CornerDirectionTransition declares dir="lu" as the default for p:strips.
    // https://dev.ecma-international.org/publications-and-standards/standards/ecma-376/
    const direction =
        effect === 'push' || effect === 'wipe'
            ? cardinalDirections.includes(rawDirection || 'l')
                ? rawDirection || 'l'
                : 'l'
            : effect === 'cover' || effect === 'pull'
              ? sideDirections.includes(rawDirection || 'l')
                  ? rawDirection || 'l'
                  : 'l'
              : effect === 'strips'
            ? stripsDirections.includes(rawDirection || 'lu')
                ? rawDirection || 'lu'
                : 'lu'
            : effect === 'split'
              ? ['in', 'out'].includes(rawDirection || 'out')
                  ? rawDirection || 'out'
                  : 'out'
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
                          : effect === 'zoom'
                            ? ['in', 'out'].includes(rawDirection || 'out')
                                ? rawDirection || 'out'
                                : 'out'
                            : rawDirection || 'l'
    if (directionSpecified && (effect === 'push' || effect === 'wipe') && !cardinalDirections.includes(rawDirection || '')) warnings.add(`The ${effect} transition direction "${rawDirection}" was normalized to l.`)
    if (directionSpecified && (effect === 'cover' || effect === 'pull') && !sideDirections.includes(rawDirection || '')) warnings.add(`The ${effect} transition direction "${rawDirection}" was normalized to l.`)
    if (effect === 'strips' && rawDirection && !stripsDirections.includes(rawDirection)) warnings.add(`The strips transition direction "${rawDirection}" was normalized to lu.`)
    if (effect === 'shred' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The shred transition direction "${rawDirection}" was normalized to in.`)
    if (effect === 'flythrough' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The flythrough transition direction "${rawDirection}" was normalized to in.`)
    if (effect === 'warp' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The warp transition direction "${rawDirection}" was normalized to out.`)
    if (effect === 'vortex' && rawDirection && !vortexDirections.includes(rawDirection)) warnings.add(`The vortex transition direction "${rawDirection}" was normalized to l.`)
    if (['prism', 'pan'].includes(effect) && rawDirection && !sideDirections.includes(rawDirection)) warnings.add(`The ${effect} transition direction "${rawDirection}" was normalized to l.`)
    if (effect === 'glitter' && rawDirection && !sideDirections.includes(rawDirection)) warnings.add(`The glitter transition direction "${rawDirection}" was normalized to l.`)
    if (['flip', 'switch', 'reveal', 'ferris', 'gallery', 'conveyor'].includes(effect) && rawDirection && !leftRightDirections.includes(rawDirection)) warnings.add(`The ${effect} transition direction "${rawDirection}" was normalized to l.`)
    if (effect === 'ripple' && rawDirection && !rippleDirections.includes(rawDirection)) warnings.add(`The ripple transition direction "${rawDirection}" was normalized to center.`)
    if (effect === 'zoom' && rawDirection && !['in', 'out'].includes(rawDirection)) warnings.add(`The zoom transition direction "${rawDirection}" was normalized to out.`)
    if (effect === 'split' && rawDirection !== undefined && !['in', 'out'].includes(rawDirection)) warnings.add(`The split transition direction "${rawDirection}" was normalized to out.`)
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
    const rawOrientation = effect === 'split' ? effectNode?.getAttribute('orient') : null
    const orientation = effect === 'split' ? (rawOrientation === 'vert' ? 'vert' : 'horz') : ['blinds', 'checker', 'comb', 'doors', 'window'].includes(effect) ? (rawDirection === 'vert' ? 'vert' : 'horz') : undefined
    if (effect === 'split' && rawOrientation !== null && rawOrientation !== 'horz' && rawOrientation !== 'vert') warnings.add(`The split transition orientation "${rawOrientation}" was normalized to horz.`)
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
    // Keep the Prism moving on the content plane and crossfade the separate slide backgrounds.
    if (isContent) warnings.add('The prism isContent flag uses a separate content prism with a background crossfade; Office rendering may differ.')
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
        ...(soundActions.length ? {soundActions} : {}),
        // MS-PPTX ST_UniversalTimeOffset has no one-minute cap; preserve valid authored durations.
        durationMs: explicitDuration !== undefined && Number.isFinite(explicitDuration) && explicitDuration >= 0 ? explicitDuration : fallbackDurationMs
    }
}


function sampleNumericAnimationFormula(
    keyframes: Array<{offset: number; value: number; formula?: string}>,
    durationMs: number,
    property: 'style.opacity' | 'fill.opacity' | 'stroke.opacity' | 'shadow.opacity' | 'style.fontSize' | 'style.fontWeight' | 'style.rotation' | 'stroke.weight' | 'shadow.offset.x' | 'shadow.offset.y' | 'imageData.cropLeft' | 'imageData.cropTop' | 'imageData.cropRight' | 'imageData.cropBottom' | 'imageData.gain' | 'imageData.blacklevel' | 'imageData.blackleve' | 'imageData.gamma' | 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' | 'ScaleX' | 'ScaleY',
    element: PptxElement,
    slideWidth: number,
    slideHeight: number,
    keyframeMode: 'lin' | 'discrete' = 'lin',
    fontWeightBaseline?: number,
    randomSeed?: number
): Array<{offset: number; value: number}> | undefined {
    const geometry = animationGeometryVariables(element, slideWidth, slideHeight)
    const isPosition = property === 'ppt_x' || property === 'ppt_y'
    const isRotation = property === 'style.rotation'
    const isImageCrop = property.startsWith('imageData.crop')
    const isImageBlacklevel = property === 'imageData.blacklevel' || property === 'imageData.blackleve'
    const isImageGain = property === 'imageData.gain'
    const isImageGamma = property === 'imageData.gamma'
    const dimensionBaseline = property === 'ppt_w' ? geometry.ppt_w : property === 'ppt_h' ? geometry.ppt_h : undefined
    const maximum = property === 'style.fontWeight' ? 1000 : property === 'style.fontSize' ? 100 : property === 'stroke.weight' ? MAX_ANIMATED_STROKE_WIDTH_EMU : property === 'shadow.offset.x' || property === 'shadow.offset.y' ? MAX_ANIMATED_SHADOW_OFFSET_EMU : isImageBlacklevel ? 0.5 : isImageGamma || isImageGain ? 32_767 : isPosition ? 100 : isRotation ? 2_147_483_647 : dimensionBaseline !== undefined ? dimensionBaseline * 10 : property === 'ScaleX' || property === 'ScaleY' ? 10 : 1
    const minimum = property === 'style.fontWeight' ? 1 : isPosition ? -100 : isRotation ? -2_147_483_647 : isImageBlacklevel ? -0.5 : isImageGain ? -32_768 : dimensionBaseline !== undefined || property === 'shadow.offset.x' || property === 'shadow.offset.y' ? -maximum : 0
    const weights = new Set(element.paragraphs.flatMap(paragraph => paragraph.runs.map(run => run.bold ? 700 : 400)))
    const variables = {
        ...geometry,
        ScaleX: 1, ScaleY: 1,
        ppt_r: element.rotation * 60_000, r: element.rotation * 60_000,
        'style.rotation': element.rotation * 60_000,
        'style.fontSize': 1,
        ...(fontWeightBaseline !== undefined
            ? {'style.fontWeight': fontWeightBaseline}
            : weights.size === 1 ? {'style.fontWeight': [...weights][0]!} : {}),
        'style.opacity': 1, 'fill.opacity': 1, 'stroke.opacity': 1, 'shadow.opacity': 1,
        'stroke.weight': element.strokeWidth,
        'shadow.offset.x': animationShadowOffsetBaseline(element, 'shadow.offset.x'),
        'shadow.offset.y': animationShadowOffsetBaseline(element, 'shadow.offset.y'),
        'imageData.cropLeft': element.imageCrop?.left ?? 0,
        'imageData.cropTop': element.imageCrop?.top ?? 0,
        'imageData.cropRight': element.imageCrop?.right ?? 0,
        'imageData.cropBottom': element.imageCrop?.bottom ?? 0,
        'imageData.gain': 1,
        'imageData.blacklevel': 0,
        'imageData.blackleve': 0,
        'imageData.gamma': 1
    }
    return sampleAnimationFormula(keyframes, durationMs, variables, minimum, maximum, keyframeMode, randomSeed)
}

function animationShadowOffsetBaseline(element: PptxElement, property: AnimationShadowOffsetProperty): number {
    const shadow = element.shadows?.[0] || element.shadow
    if (!shadow) return 0
    const angle = (shadow.angle || 0) * Math.PI / 180
    return (property === 'shadow.offset.x' ? Math.cos(angle) : Math.sin(angle)) * shadow.distance
}

function supportsStrokeOpacity(element: PptxElement): boolean {
    if (element.kind === 'shape' && ['cube', 'plaque', 'frame', 'halfFrame', 'cornerTabs', 'plaqueTabs', 'squareTabs', 'donut', 'pie', 'chord', 'arc', 'blockArc', 'pieWedge'].includes(element.geometry)) return true
    return ['shape', 'line'].includes(element.kind) && (!/gradient\(/i.test(element.stroke)
        || element.kind === 'shape' && (['mathDivide', 'mathMultiply', 'mathNotEqual', 'round1Rect', 'round2DiagRect', 'round2SameRect', 'snipRoundRect'].includes(element.geometry) || ['rect', 'ellipse', 'roundRect', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDelay', 'flowChartDisplay', 'flowChartOfflineStorage', 'flowChartOnlineStorage', 'flowChartPunchedTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartMagneticTape', 'flowChartMultidocument', 'flowChartInternalStorage', 'flowChartSummingJunction', 'flowChartOr', 'plus', 'mathMinus', 'mathEqual', 'mathPlus', 'snip1Rect', 'snip2DiagRect', 'snip2SameRect', 'homePlate', 'nonIsoscelesTrapezoid', 'heptagon', 'decagon', 'leftRightArrow', 'upDownArrow', 'upArrow', 'downArrow', 'bentUpArrow', 'quadArrow', 'leftRightUpArrow', 'rightArrow', 'leftArrow', 'notchedRightArrow', 'heart', 'lightningBolt', 'triangle', 'rtTriangle', 'diamond', 'flowChartConnector', 'flowChartDecision', 'flowChartSort', 'flowChartCollate', 'flowChartPunchedCard', 'flowChartOffpageConnector', 'flowChartDocument', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartInputOutput', 'flowChartPreparation', 'flowChartTerminator', 'flowChartPredefinedProcess', 'flowChartExtract', 'flowChartMerge', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32'].includes(element.geometry || '') || element.geometry === 'custom' && Boolean(element.customPaths?.length)))
}

function hasOnOffBaseline(element: PptxElement, property: AnimationOnOffProperty): boolean {
    if (property === 'fill.on') return element.kind === 'shape' && element.fill !== 'transparent'
    if (property === 'stroke.on') return ['shape', 'line'].includes(element.kind) && element.stroke !== 'transparent' && supportsStrokeOpacity(element)
    return hasShadow(element)
}

function hasShadow(element: PptxElement): boolean {
    return Boolean(element.shadow || element.shadows?.length || element.paragraphs.some(paragraph => paragraph.runs.some(run => run.shadow || run.shadows?.length)))
}

function parseOnOffValue(value: string | null): boolean | undefined {
    const normalized = value?.trim().toLowerCase()
    return normalized === 'true' || normalized === 't' ? true : normalized === 'false' || normalized === 'f' ? false : undefined
}

function parseStrokeDashStyle(value: string | null): AnimatedStrokeDashStyle | undefined {
    const styles: Record<string, AnimatedStrokeDashStyle> = {
        solid: 'solid',
        dot: 'dot',
        dash: 'dash',
        dashdot: 'dashDot',
        longdash: 'lgDash',
        longdashdot: 'lgDashDot',
        longdashdotdot: 'lgDashDotDot'
    }
    return value ? styles[value.trim().toLowerCase()] : undefined
}

function parseLineArrow(value: string | null): AnimatedLineArrow | undefined {
    const normalized = value?.trim().toLowerCase()
    return normalized === 'none' || normalized === 'block' || normalized === 'classic' || normalized === 'diamond' || normalized === 'oval' || normalized === 'open' || normalized === 'chevron' || normalized === 'doublechevron'
        ? normalized
        : undefined
}

function parseLineArrowSize(property: AnimatedLineArrowSizeProperty, value: string | null): AnimatedLineArrowSize | undefined {
    const normalized = value?.trim().toLowerCase()
    // MS-OE376 defines width and length as discrete string presets, with separate values for each attribute.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7b427ccc-3a1f-418b-821f-d78f7aed51c5
    const allowed = property.endsWith('Width') ? ['narrow', 'medium', 'wide'] : ['short', 'medium', 'long']
    return normalized && allowed.includes(normalized) ? normalized as AnimatedLineArrowSize : undefined
}

type RandomEffectVariant = Pick<PptxAnimation, 'effect' | 'wipeDirection' | 'blindsOrientation' | 'checkerOrientation' | 'randomBarOrientation' | 'stripsDirection' | 'wheelSpokes' | 'scaleFrom' | 'scaleTo'>

// PowerPoint and WPS can serialize preset 24 as an empty p:anim without recording the selected visual effect.
// Choose one supported variant per import; the parsed action keeps that choice for the session's replays.
const RANDOM_EFFECT_VARIANTS: RandomEffectVariant[] = [
    {effect: 'fade'},
    {effect: 'wipe', wipeDirection: 'left'},
    {effect: 'blinds', blindsOrientation: 'horizontal'},
    {effect: 'checker', checkerOrientation: 'horz'},
    {effect: 'randomBars', randomBarOrientation: 'horizontal'},
    {effect: 'strips', stripsDirection: 'rd'},
    {effect: 'dissolve'},
    {effect: 'wheel', wheelSpokes: 4},
    {effect: 'zoom', scaleFrom: [0, 0], scaleTo: [1, 1]}
]

function randomEffectVariant(): RandomEffectVariant {
    return RANDOM_EFFECT_VARIANTS[Math.floor(Math.random() * RANDOM_EFFECT_VARIANTS.length)]!
}

async function parseAnimations(slideRoot: Element, slidePath: string, theme: Record<string, string>, warnings: Set<string>, elements: PptxElement[], slideWidth: number, slideHeight: number, smartArtAnimationTargets: Map<string, string>, rels: Map<string, PptxRelationship>, zip: JSZip, urls: string[], assetUrlCache: Map<string, string>, isWpsApplication: boolean): Promise<Pick<PptxSlide, 'automaticAnimations' | 'animationSteps' | 'animationSequence' | 'animationStepSequences' | 'triggeredAnimations'>> {
    // p:timing is a direct p:sld child; ignore same-named nodes embedded in extension payloads.
    // Source: https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-presentation-slides
    const timing = child(slideRoot, 'timing')
    if (!timing) return {}
    const soundTargetUrls = new Map<Element, string>()
    const soundTargetIds = new Map<Element, string>()
    let soundTargetIndex = 0
    for (const mediaNode of descendants(timing, 'audio').concat(descendants(timing, 'video')).filter(node => node.namespaceURI === PRESENTATION_NS)) {
        const soundTarget = child(child(child(mediaNode, 'cMediaNode'), 'tgtEl'), 'sndTgt')
        if (!soundTarget || soundTarget.namespaceURI !== PRESENTATION_NS) continue
        const relationshipId = namespacedAttr(soundTarget, 'embed')
        const linkId = namespacedAttr(soundTarget, 'link')
        const relationship = relationshipId ? rels.get(relationshipId) : undefined
        if (!relationshipId || !relationship || relationship.external) {
            warnings.add(linkId && rels.get(linkId)?.external
                ? 'An external animation sound target was skipped; only embedded audio is supported.'
                : 'An animation sound target without a resolvable embedded audio relationship was skipped.')
            continue
        }
        const url = await transitionAudioUrlFor(relationship.target, zip, urls, assetUrlCache)
        if (!url) {
            warnings.add('An animation sound target did not resolve to a supported embedded audio file.')
            continue
        }
        const timeNode = child(child(mediaNode, 'cMediaNode'), 'cTn')
        const timeNodeId = timeNode?.getAttribute('id')
        soundTargetUrls.set(mediaNode, url)
        soundTargetIds.set(mediaNode, `pptx-sound-target-${timeNodeId || ++soundTargetIndex}`)
    }
    const isInSubTimeNodeList = (node: Element) => {
        let parent = node.parentElement
        while (parent && parent !== timing) {
            if (parent.namespaceURI === PRESENTATION_NS && parent.localName === 'subTnLst') return true
            parent = parent.parentElement
        }
        return false
    }
    const handledSubTimeEntries = new Set<Element>()
    let exclusiveGroupId = 0
    let subTimeFillScopeId = 0
    const subTimeLeafTypes = new Set(['anim', 'animClr', 'animEffect', 'animMotion', 'animRot', 'animScale', 'set', 'cmd', 'audio', 'video'])
    const subTimeContainerTypes = new Set(['par', 'seq', 'excl'])
    const subTimeEntryTypes = new Set([...subTimeLeafTypes, ...subTimeContainerTypes])
    const supportsParallelSubTimeTiming = (timeNode: Element | undefined) => {
        const entries = timeNode ? children(child(timeNode, 'childTnLst')) : []
        return timeNode?.parentElement?.localName === 'par' && entries.length > 0 && entries.every(entry =>
            subTimeLeafTypes.has(entry.localName) && entry.localName !== 'cmd'
        )
    }
    const supportsFlattenedParallelSubTimeTiming = (timeNode: Element | undefined, actions: PptxAnimation[]) => {
        const entries = timeNode ? children(child(timeNode, 'childTnLst')) : []
        const groupTimingAttributes = ['accel', 'decel', 'tmFilter', 'spd', 'repeatCount', 'repeatDur', 'autoRev']
        const canFlattenWithoutNestedClock = (entry: Element): boolean => {
            if (subTimeLeafTypes.has(entry.localName)) return entry.localName !== 'cmd'
            if (!['par', 'seq'].includes(entry.localName)) return false
            const nestedTimeNode = child(entry, 'cTn')
            const nestedEntries = children(child(nestedTimeNode, 'childTnLst'))
            const nestedSequenceControls = entry.localName === 'seq' && (
                Boolean(child(entry, 'nextCondLst') || child(entry, 'prevCondLst'))
                || (entry.hasAttribute('concurrent') && !['0', 'false'].includes(entry.getAttribute('concurrent') || ''))
                || (entry.hasAttribute('nextAc') && entry.getAttribute('nextAc') !== 'none')
                || (entry.hasAttribute('prevAc') && entry.getAttribute('prevAc') !== 'none')
            )
            return Boolean(nestedTimeNode && nestedEntries.length
                && !groupTimingAttributes.some(attribute => nestedTimeNode.hasAttribute(attribute))
                && !child(nestedTimeNode, 'endCondLst')
                && !child(nestedTimeNode, 'iterate')
                && !child(nestedTimeNode, 'subTnLst')
                && !nestedTimeNode.hasAttribute('masterRel')
                && !nestedTimeNode.hasAttribute('afterEffect')
                && !nestedSequenceControls
                && nestedEntries.every(canFlattenWithoutNestedClock))
        }
        // A parallel time node can carry its clock through simple nested groups when none of them introduces another clock.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.paralleltimenode?view=openxml-3.0.1
        return timeNode?.parentElement?.localName === 'par'
            && entries.length > 0
            && entries.every(entry => ['par', 'seq'].includes(entry.localName))
            && entries.every(canFlattenWithoutNestedClock)
            && !child(timeNode, 'iterate')
            && !child(timeNode, 'subTnLst')
            && !timeNode.hasAttribute('afterEffect')
            && actions.length > 0
            && actions.every(action => !action.timingWarp && (!action.mediaWaitForEnd || Number.isFinite(action.parentEndMs)) && !action.mediaWaitForEndKeys?.length)
    }
    const supportsSequentialSubTimeTiming = (timeNode: Element | undefined, actions: PptxAnimation[]) => {
        const entries = timeNode ? children(child(timeNode, 'childTnLst')) : []
        const groupTimingAttributes = ['accel', 'decel', 'tmFilter', 'spd', 'repeatCount', 'repeatDur', 'autoRev']
        const canFlattenWithoutNestedClock = (entry: Element): boolean => {
            if (subTimeLeafTypes.has(entry.localName)) return entry.localName !== 'cmd'
            if (!['par', 'seq'].includes(entry.localName)) return false
            const nestedTimeNode = child(entry, 'cTn')
            const nestedEntries = children(child(nestedTimeNode, 'childTnLst'))
            const nestedSequenceControls = entry.localName === 'seq' && (
                Boolean(child(entry, 'nextCondLst') || child(entry, 'prevCondLst'))
                || (entry.hasAttribute('concurrent') && !['0', 'false'].includes(entry.getAttribute('concurrent') || ''))
                || (entry.hasAttribute('nextAc') && entry.getAttribute('nextAc') !== 'none')
                || (entry.hasAttribute('prevAc') && entry.getAttribute('prevAc') !== 'none')
            )
            return Boolean(nestedTimeNode && nestedEntries.length
                && !groupTimingAttributes.some(attribute => nestedTimeNode.hasAttribute(attribute))
                && !child(nestedTimeNode, 'endCondLst')
                && !child(nestedTimeNode, 'iterate')
                && !child(nestedTimeNode, 'subTnLst')
                && !nestedTimeNode.hasAttribute('masterRel')
                && !nestedTimeNode.hasAttribute('afterEffect')
                && !nestedSequenceControls
                && nestedEntries.every(canFlattenWithoutNestedClock))
        }
        // A sequential time node waits for each preceding child. Apply its common-time speed to flattened nested groups only when none defines a second clock.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.sequencetimenode?view=openxml-3.0.1
        return timeNode?.parentElement?.localName === 'seq'
            && entries.length > 0
            && entries.every(canFlattenWithoutNestedClock)
            && actions.length > 0
            && actions.every(action => !action.timingWarp && (!action.mediaWaitForEnd || Number.isFinite(action.parentEndMs)) && !action.mediaWaitForEndKeys?.length)
    }
    const subTimeEntryTimeNode = (entry: Element) => ['audio', 'video'].includes(entry.localName)
        ? child(child(entry, 'cMediaNode'), 'cTn')
        : child(child(entry, 'cBhvr'), 'cTn') || child(entry, 'cTn')
    const MAX_TL_TIME_MS = 2_147_483_625
    // OOXML timing node kinds and delay semantics:
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timenodevalues
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition
    // ECMA-376 ST_TLTime is an unsigned integer in milliseconds or `indefinite`.
    // https://ecma-international.org/publications-and-standards/standards/ecma-376/
    // OOXML simple types use XML whitespace, so trim only space, tab, CR, and LF—not JS's wider Unicode trim set.
    const isFiniteTimelineTime = (value: string | null | undefined) => {
        const normalized = trimXmlWhitespace(value)
        return normalized !== undefined && /^\+?\d+$/.test(normalized)
    }
    const isIndefiniteTimelineTime = (value: string | null | undefined) => trimXmlWhitespace(value) === 'indefinite'
    const timelineTime = (value: string | null | undefined, fallback = 0) => {
        const normalized = trimXmlWhitespace(value)
        if (!normalized || normalized === 'indefinite' || !isFiniteTimelineTime(normalized)) return fallback
        const parsed = Number(normalized)
        if (!Number.isFinite(parsed)) return fallback
        if (parsed > MAX_TL_TIME_MS) {
            warnings.add(`A timing value exceeds the PowerPoint ST_TLTime maximum and was capped at ${MAX_TL_TIME_MS}.`)
            return MAX_TL_TIME_MS
        }
        return parsed
    }
    const deferredEndConditionWarnings = new Map<Element, string>()
    const offsetEndDuration = (timeNode: Element | undefined, effect: string): number | undefined => {
        const endConditionList = child(timeNode, 'endCondLst')
        if (!endConditionList) return undefined
        const conditions = children(endConditionList, 'cond')
        if (!conditions.length) {
            warnings.add(`${effect} has an empty end-condition list; its declared duration was used.`)
            return undefined
        }
        const delays: number[] = []
        for (const condition of conditions) {
            if (Array.from(condition.attributes).some(attribute => attribute.localName !== 'delay') || condition.children.length) {
                if (timeNode && (condition.hasAttribute('evt') || condition.children.length)) deferredEndConditionWarnings.set(timeNode, effect)
                else warnings.add(`${effect} has an invalid end condition; its declared duration was used.`)
                return undefined
            }
            const delay = condition.getAttribute('delay')
            if (isIndefiniteTimelineTime(delay)) {
                delays.push(Infinity)
                continue
            }
            if (delay !== null && !isFiniteTimelineTime(delay)) {
                warnings.add(`${effect} has an invalid end-condition delay; its declared duration was used.`)
                return undefined
            }
            delays.push(delay === null ? 0 : timelineTime(delay, Number.NaN))
        }
        if (delays.some(Number.isNaN)) {
            warnings.add(`${effect} has an invalid end-condition delay; its declared duration was used.`)
            return undefined
        }
        // OOXML defines the list as conditions that must be met to stop the node.
        // Offset-only conditions therefore stop it at the latest listed delay.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.endconditionlist?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1
        return Math.max(...delays)
    }
    const subTimeConditionEventTime = (
        timeNode: Element,
        condition: Element,
        sources: Map<string, {startMs: number; endMs?: number}>
    ): {eventAtMs: number; delayMs: number} | undefined => {
        if (Array.from(condition.attributes).some(attribute => !['evt', 'delay'].includes(attribute.localName))
            || condition.children.length !== 1
            || condition.children[0]?.namespaceURI !== PRESENTATION_NS) return undefined
        const event = condition.getAttribute('evt')
        const trigger = condition.children[0]
        if (event !== 'begin' && event !== 'end') return undefined
        const rawDelay = condition.getAttribute('delay')
        if (isIndefiniteTimelineTime(rawDelay) || (rawDelay !== null && !isFiniteTimelineTime(rawDelay))) return undefined
        const delayMs = rawDelay === null ? 0 : timelineTime(rawDelay, Number.NaN)
        if (!Number.isFinite(delayMs)) return undefined
        if (trigger.localName === 'tn') {
            const targetId = trigger.getAttribute('val')
            const source = targetId ? sources.get(targetId) : undefined
            const eventAtMs = event === 'begin' ? source?.startMs : source?.endMs
            return Number.isFinite(eventAtMs) ? {eventAtMs: eventAtMs!, delayMs} : undefined
        }
        if (trigger.localName !== 'rtn') return undefined
        const triggerValue = trigger.getAttribute('val')
        if (!['first', 'last', 'all'].includes(triggerValue || '')) return undefined
        const childTimeNodes = children(child(timeNode, 'childTnLst')).map(node => child(node, 'cTn'))
        if (!childTimeNodes.length) return undefined
        const childTimes = childTimeNodes.map(childTimeNode => {
            const id = childTimeNode?.getAttribute('id')
            return id ? sources.get(id) : undefined
        })
        if (childTimes.some(source => !source || !Number.isFinite(source.startMs) || !Number.isFinite(source.endMs))) return undefined
        const resolvedTimes = childTimes as Array<{startMs: number; endMs: number}>
        const eventTime = (source: {startMs: number; endMs: number}) => event === 'begin' ? source.startMs : source.endMs
        let eventAtMs: number
        if (triggerValue === 'all') {
            eventAtMs = Math.max(...resolvedTimes.map(eventTime))
        } else {
            // Runtime-node order is based on child end time, even when the trigger event is begin.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.runtimenodetrigger?view=openxml-3.0.1
            const ordered = [...resolvedTimes].sort((left, right) => left.endMs - right.endMs)
            eventAtMs = eventTime(triggerValue === 'first' ? ordered[0]! : ordered[ordered.length - 1]!)
        }
        return {eventAtMs, delayMs}
    }
    const subTimeEventEndDuration = (
        timeNode: Element,
        timeNodeStartMs: number,
        sources: Map<string, {startMs: number; endMs?: number}>
    ): number | undefined => {
        const conditions = children(child(timeNode, 'endCondLst'), 'cond')
        if (!conditions.length) return undefined
        const endTimes: number[] = []
        for (const condition of conditions) {
            if (Array.from(condition.attributes).some(attribute => !['evt', 'delay'].includes(attribute.localName))) return undefined
            const rawDelay = condition.getAttribute('delay')
            if (isIndefiniteTimelineTime(rawDelay) || (rawDelay !== null && !isFiniteTimelineTime(rawDelay))) return undefined
            const delayMs = rawDelay === null ? 0 : timelineTime(rawDelay, Number.NaN)
            if (!Number.isFinite(delayMs)) return undefined
            const event = condition.getAttribute('evt')
            let eventAtMs: number | undefined
            if (event === null && condition.children.length === 0) {
                eventAtMs = timeNodeStartMs
            } else if ((event === 'begin' || event === 'end')
                && condition.children.length === 1
                && condition.children[0]?.namespaceURI === PRESENTATION_NS
                && condition.children[0]?.localName === 'tn') {
                const targetId = condition.children[0].getAttribute('val')
                const source = targetId ? sources.get(targetId) : undefined
                eventAtMs = event === 'begin' ? source?.startMs : source?.endMs
            } else if ((event === 'begin' || event === 'end')
                && condition.children.length === 1
                && condition.children[0]?.namespaceURI === PRESENTATION_NS
                && condition.children[0]?.localName === 'rtn') {
                eventAtMs = subTimeConditionEventTime(timeNode, condition, sources)?.eventAtMs
            } else {
                return undefined
            }
            if (!Number.isFinite(eventAtMs)) return undefined
            endTimes.push(eventAtMs! + delayMs)
        }
        // A condition list is conjunctive, so its stop time is the latest satisfied condition.
        // Microsoft defines begin/end as events raised by the referenced time node.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/6acf6d2c-6423-4f5e-8a16-6533a20fbd41
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1
        return Math.max(0, Math.max(...endTimes) - timeNodeStartMs)
    }
    const subTimeEventStartTime = (
        timeNode: Element,
        parentStartMs: number,
        sources: Map<string, {startMs: number; endMs?: number}>
    ): number | undefined => {
        const conditions = children(child(timeNode, 'stCondLst'), 'cond')
        if (!conditions.length) return parentStartMs
        const startTimes: number[] = []
        for (const condition of conditions) {
            if (Array.from(condition.attributes).some(attribute => !['evt', 'delay'].includes(attribute.localName))) return undefined
            const rawDelay = condition.getAttribute('delay')
            if (isIndefiniteTimelineTime(rawDelay) || (rawDelay !== null && !isFiniteTimelineTime(rawDelay))) return undefined
            const delayMs = rawDelay === null ? 0 : timelineTime(rawDelay, Number.NaN)
            if (!Number.isFinite(delayMs)) return undefined
            const event = condition.getAttribute('evt')
            let eventAtMs: number | undefined
            if (event === null && condition.children.length === 0) {
                eventAtMs = parentStartMs
            } else if ((event === 'begin' || event === 'end')
                && condition.children.length === 1
                && condition.children[0]?.namespaceURI === PRESENTATION_NS
                && condition.children[0]?.localName === 'tn') {
                const targetId = condition.children[0].getAttribute('val')
                const source = targetId ? sources.get(targetId) : undefined
                eventAtMs = event === 'begin' ? source?.startMs : source?.endMs
            } else if ((event === 'begin' || event === 'end')
                && condition.children.length === 1
                && condition.children[0]?.namespaceURI === PRESENTATION_NS
                && condition.children[0]?.localName === 'rtn') {
                eventAtMs = subTimeConditionEventTime(timeNode, condition, sources)?.eventAtMs
            } else {
                return undefined
            }
            if (!Number.isFinite(eventAtMs)) return undefined
            startTimes.push(eventAtMs! + delayMs)
        }
        // Each start condition must be met before the time node activates.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.startconditionlist?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1
        return Math.max(parentStartMs, ...startTimes)
    }
    const animationTiming = (timeNode: Element | undefined, effect: string, resolvedEndConditionDurationMs?: number): Pick<PptxAnimation, 'sourceTimeNodeId' | 'fillMode' | 'repeatCount' | 'repeatDurationMs' | 'autoReverse' | 'acceleration' | 'deceleration' | 'bounceEndRatio' | 'timeFilter' | 'playbackSpeed' | 'reversePlayback'> | undefined => {
        const rawFill = timeNode?.getAttribute('fill')
        const fill = rawFill === null || rawFill === undefined ? 'hold' : trimXmlWhitespace(rawFill)
        const rawAutoRev = timeNode?.getAttribute('autoRev')
        const autoRev = rawAutoRev === null || rawAutoRev === undefined ? undefined : trimXmlWhitespace(rawAutoRev)
        // fill is an ST_TLTimeNodeFillType value and autoRev is an XML boolean; explicit empty values are invalid, not omitted defaults.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/9bdaa76e-c714-4034-8c29-ac42071e3484
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
        if (!fill || !['hold', 'freeze', 'remove', 'transition'].includes(fill) || (autoRev !== undefined && !['0', '1', 'false', 'true'].includes(autoRev))) {
            warnings.add(`${effect} animations with unsupported fill or auto-reverse timing are not rendered yet.`)
            return undefined
        }
        const countValue = timeNode?.getAttribute('repeatCount')
        const normalizedCountValue = trimXmlWhitespace(countValue)
        const durationValue = timeNode?.getAttribute('repeatDur')
        // ST_TLTime is unsignedInt or `indefinite`; reject decimal values instead of treating them as fractional iterations.
        const parsedCount = normalizedCountValue === undefined || isIndefiniteTimelineTime(normalizedCountValue)
            ? undefined
            : isFiniteTimelineTime(normalizedCountValue)
                ? timelineTime(normalizedCountValue, Number.NaN) / 1000
                : Number.NaN
        const count = normalizedCountValue === undefined
            ? (durationValue === null || durationValue === undefined ? undefined : Infinity)
            : isIndefiniteTimelineTime(normalizedCountValue)
                ? Infinity
                : parsedCount !== undefined && parsedCount >= 0 && parsedCount < 1
                    ? Infinity
                    : parsedCount
        const rawSpeed = timeNode?.getAttribute('spd')
        const speed = parseAnimationSpeed(rawSpeed)
        // spd is a timing-rate percentage; retain zero so the local timeline stays at its start position.
        // Microsoft defines a negative cTn@spd as reversed timing; use its magnitude for rate and retain its sign for direction.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
        const playbackSpeed = speed ?? 1
        if (speed === undefined) warnings.add(`${effect} animations with invalid speed are played at normal speed.`)
        const playbackRate = Math.abs(playbackSpeed)
        const authoredRepeatDurationMs = (() => {
            if (durationValue === null || durationValue === undefined) return undefined
            if (isIndefiniteTimelineTime(durationValue)) return Infinity
            const repeatDurationMs = timelineTime(durationValue, Number.NaN)
            if (playbackRate === 0) return repeatDurationMs === 0 ? 0 : Infinity
            return repeatDurationMs / playbackRate
        })()
        const endConditionDurationMs = resolvedEndConditionDurationMs ?? offsetEndDuration(timeNode, effect)
        const repeatDurationMs = endConditionDurationMs === undefined
            ? authoredRepeatDurationMs
            : Math.min(authoredRepeatDurationMs ?? Infinity, endConditionDurationMs)
        if ((count !== undefined && !(count > 0)) || (repeatDurationMs !== undefined && (Number.isNaN(repeatDurationMs) || repeatDurationMs < 0))) {
            warnings.add(`${effect} animations with invalid repeat timing are not rendered yet.`)
            return undefined
        }
        const fixedPercentage = (value: string | null): number | undefined => {
            if (value === null) return 0
            const percent = percentageValue(value)
            return percent !== undefined && percent >= 0 && percent <= 1 ? percent : undefined
        }
        const acceleration = fixedPercentage(timeNode?.getAttribute('accel') ?? null)
        const deceleration = fixedPercentage(timeNode?.getAttribute('decel') ?? null)
        if (acceleration === undefined || deceleration === undefined || acceleration + deceleration > 1) {
            warnings.add(`${effect} animations with invalid acceleration/deceleration timing are not rendered yet.`)
            return undefined
        }
        const rawTimeFilter = timeNode?.getAttribute('tmFilter')
        const timeFilter = parseTimingFilter(rawTimeFilter)
        if (rawTimeFilter?.trim() && !timeFilter) warnings.add(`${effect} animations with an unsupported time filter use their unfiltered timing.`)
        let behavior: Element | null | undefined = timeNode
        while (behavior && !['anim', 'animMotion', 'animRot', 'animScale'].includes(behavior.localName)) behavior = behavior.parentElement
        const hasBounceEnd = behavior?.hasAttributeNS(POWERPOINT_2010_NS, 'bounceEnd') || false
        const rawBounceEnd = hasBounceEnd ? behavior?.getAttributeNS(POWERPOINT_2010_NS, 'bounceEnd') ?? null : null
        const bounceEndRatio = hasBounceEnd ? fixedPercentage(rawBounceEnd) : undefined
        if (hasBounceEnd && (bounceEndRatio === undefined || bounceEndRatio <= 0 || bounceEndRatio >= 1)) {
            warnings.add(`${effect} has an invalid or unsupported p14:bounceEnd value; bounce timing was ignored.`)
        } else if (bounceEndRatio !== undefined && bounceEndRatio > 0) {
            warnings.add(`${effect} p14:bounceEnd timing uses the player's damped-spring approximation; Office's exact spring parameters are not specified.`)
        }
        // OOXML repeatCount uses thousandths of an iteration. Microsoft says PowerPoint does not complete nodes below 1000;
        // interpret those counts as active until repeatDur or a parent end bounds them.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/02b2312e-9700-42e0-8a19-b2af1ba83a20
        // autoRev doubles each simple-duration cycle.
        return {
            sourceTimeNodeId: timeNode?.getAttribute('id') || undefined,
            fillMode: fill === 'remove' || fill === 'transition' ? 'remove' : (fill as 'hold' | 'freeze'),
            repeatCount: count,
            repeatDurationMs,
            autoReverse: autoRev === '1' || autoRev === 'true',
            acceleration,
            deceleration,
            bounceEndRatio: bounceEndRatio !== undefined && bounceEndRatio > 0 && bounceEndRatio < 1 ? bounceEndRatio : undefined,
            timeFilter,
            playbackSpeed,
            reversePlayback: playbackSpeed < 0
        }
    }
    const parentFillEnd = (action: PptxAnimation, parentTimeNode: Element | undefined, endMs: number): Pick<PptxAnimation, 'parentFillEndMs'> => {
        const rawFill = parentTimeNode?.getAttribute('fill')
        // Child hold/freeze values survive while their parent is running or holding.
        // Only a removing/transitioning parent makes its end a reset boundary.
        const fill = rawFill === null || rawFill === undefined ? 'hold' : trimXmlWhitespace(rawFill)
        if ((fill !== 'remove' && fill !== 'transition') || !Number.isFinite(endMs)) return {}
        return {parentFillEndMs: Math.min(action.parentFillEndMs ?? Infinity, endMs)}
    }
    const animationDurationMs = (timeNode: Element | undefined, timing: Pick<PptxAnimation, 'playbackSpeed'>, fallback = 500) => {
        const duration = timeNode?.getAttribute('dur')
        // ST_TLTime permits an indefinite duration; preserve it as an unending animation.
        // An omitted cTn@dur is a different case and PowerPoint defaults it to 0.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
        // PowerPoint defaults an omitted cTn@dur to 0; keep the existing fallback for invalid values.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/02b2312e-9700-42e0-8a19-b2af1ba83a20
        if (isIndefiniteTimelineTime(duration)) return Infinity
        const durationMs = timelineTime(duration, duration === null || duration === undefined ? 0 : fallback)
        if (timing.playbackSpeed === 0) return durationMs === 0 ? 0 : Infinity
        return durationMs / Math.abs(timing.playbackSpeed ?? 1)
    }
    // p:cMediaNode carries a p:cTn. When the media duration is unknown, a finite parent end bounds the
    // playback interval; treating the natural media end as unbounded here is an implementation inference.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/9fb1b81d-d5a1-4aae-911f-d0be67f559f3
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
    const basePlaybackDuration = (action: PptxAnimation) =>
        action.effect === 'media' ? (action.mediaDurationMs ?? (action.mediaWaitForEnd ? Infinity : 0)) : action.durationMs === 0 ? 0 : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (action.repeatCount ?? 1), action.repeatDurationMs ?? Infinity)
    const iterationSpan = (action: PptxAnimation) => action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0
    const playbackDuration = (action: PptxAnimation) => Math.min(
        action.timingWarp
            ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs)
            : basePlaybackDuration(action) + iterationSpan(action),
        action.parentEndMs === undefined ? Infinity : Math.max(0, action.parentEndMs - action.delayMs)
    )
    type GroupTiming = TimingCurve & {durationMs: number; playbackSpeed: number; playbackDurationMs: number; reversePlayback: boolean; autoReverse: boolean}
    const groupTimingFor = (owner: Element, effect: string, childDerivedDurationMs?: number, resolvedEndConditionDurationMs?: number): GroupTiming | undefined => {
        const timing = animationTiming(owner, effect, resolvedEndConditionDurationMs)
        if (!timing) return undefined
        const groupSpeed = Math.abs(timing.playbackSpeed ?? 1)
        const rawDuration = owner.getAttribute('dur')
        const declaredDurationMs = timelineTime(rawDuration, Number.NaN)
        // PowerPoint defaults omitted cTn@dur to 0 and derives a zero-duration time node from its children.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/02b2312e-9700-42e0-8a19-b2af1ba83a20
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/13b03cf8-e193-4d7c-bece-efee91ea40bd
        // A finite end condition can bound a container whose simple duration is indefinite.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.endconditionlist?view=openxml-3.0.1
        const durationMs = rawDuration === null || declaredDurationMs === 0
            || isIndefiniteTimelineTime(rawDuration) && Number.isFinite(timing.repeatDurationMs)
            ? childDerivedDurationMs ?? Number.NaN
            : declaredDurationMs
        if (!Number.isFinite(durationMs) || durationMs <= 0) {
            warnings.add('A time-container timing curve or repeat without a finite positive duration is not applied.')
            return undefined
        }
        const cycleDurationMs = groupSpeed === 0 ? Infinity : (durationMs * (timing.autoReverse ? 2 : 1)) / groupSpeed
        const repeatCount = timing.repeatCount ?? (owner.hasAttribute('repeatDur') ? Infinity : 1)
        const playbackDurationMs = Math.min(cycleDurationMs * repeatCount, timing.repeatDurationMs ?? Infinity)
        if (Number.isNaN(playbackDurationMs) || playbackDurationMs <= 0) {
            warnings.add('An indefinite or empty parent time-container repeat is not applied yet.')
            return undefined
        }
        return {
            durationMs,
            playbackSpeed: groupSpeed,
            playbackDurationMs,
            reversePlayback: timing.reversePlayback ?? false,
            autoReverse: timing.autoReverse ?? false,
            acceleration: timing.acceleration ?? 0,
            deceleration: timing.deceleration ?? 0,
            timeFilter: timing.timeFilter
        }
    }
    const applyGroupTiming = (actions: PptxAnimation[], delayMs: number, groupTiming: GroupTiming, perIteration = false, parentTimeNode?: Element) => actions.map(action => {
        if (action.effect === 'media') {
            const childOffsetMs = action.delayMs - delayMs
            const cycleDurationMs = groupTiming.durationMs / groupTiming.playbackSpeed
            // cTn@spd scales the child timeline. A single constant-rate parent cycle maps directly
            // to a media playback rate and duration; acceleration, reversal, and repeats need a
            // continuously controlled media clock and stay unsupported here.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
            // p:audio/p:video carry their own cTn through cMediaNode.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonmedianode?view=openxml-3.0.1
            const hasSingleConstantRateCycle = groupTiming.playbackSpeed > 0
                && !groupTiming.reversePlayback
                && !groupTiming.autoReverse
                && groupTiming.acceleration === 0
                && groupTiming.deceleration === 0
                && !groupTiming.timeFilter
                && !perIteration
                && !action.iteration
                && !action.timingWarp
                && Number.isFinite(cycleDurationMs)
                && Number.isFinite(groupTiming.playbackDurationMs)
                && Math.abs(groupTiming.playbackDurationMs - cycleDurationMs) < 0.01
            if (!hasSingleConstantRateCycle
                || !Number.isFinite(childOffsetMs)
                || childOffsetMs < 0) {
                warnings.add('A time-container timing curve or repeat on timed media is not applied yet.')
                return action
            }
            const isPlaybackCommand = action.mediaCommand === 'play' || action.mediaCommand === 'togglePause'
            const groupEndDelayMs = delayMs + groupTiming.playbackDurationMs
            return {
                ...action,
                delayMs: delayMs + childOffsetMs / groupTiming.playbackSpeed,
                ...(isPlaybackCommand ? {
                    mediaPlaybackRate: (action.mediaPlaybackRate ?? 1) * groupTiming.playbackSpeed,
                    ...(action.mediaDurationMs !== undefined
                        ? {mediaDurationMs: action.mediaDurationMs / groupTiming.playbackSpeed}
                        : {})
                } : {}),
                parentEndMs: Math.min(action.parentEndMs ?? Infinity, groupEndDelayMs),
                ...parentFillEnd(action, parentTimeNode, groupEndDelayMs),
                timingWarp: {
                    acceleration: groupTiming.acceleration,
                    deceleration: groupTiming.deceleration,
                    timeFilter: groupTiming.timeFilter,
                    groupOffsetMs: delayMs,
                    groupDurationMs: groupTiming.durationMs,
                    groupSpeed: groupTiming.playbackSpeed,
                    groupReversePlayback: groupTiming.reversePlayback,
                    groupAutoReverse: groupTiming.autoReverse,
                    groupPlaybackDurationMs: groupTiming.playbackDurationMs,
                    groupPerIteration: false,
                    childOffsetMs,
                    groupEndDelayMs
                }
            }
        }
        const childOffsetMs = action.delayMs - delayMs
        const appliesPerIteration = perIteration && Boolean(action.iteration)
        const childNeverAdvances = action.playbackSpeed === 0 && action.durationMs > 0
        const childDurationMs = childNeverAdvances
            ? Math.max(0, groupTiming.durationMs - childOffsetMs)
            : appliesPerIteration ? basePlaybackDuration(action) : playbackDuration(action)
        if (!Number.isFinite(childOffsetMs) || childOffsetMs < 0 || !Number.isFinite(childDurationMs) || childOffsetMs + childDurationMs > groupTiming.durationMs + 0.01) {
            warnings.add('A time-container timing curve or repeat is not applied because child timing exceeds its finite duration.')
            return action
        }
        const firstChildOffsetMs = groupTiming.reversePlayback ? childOffsetMs + childDurationMs : childOffsetMs
        let wallStartMs: number
        if (groupTiming.playbackSpeed === 0) {
            const frozenGroupTimeMs = timingProgress(groupTiming, 0) * groupTiming.durationMs
            wallStartMs = firstChildOffsetMs <= frozenGroupTimeMs ? 0 : Infinity
        } else {
            const scale = groupTiming.durationMs / groupTiming.playbackSpeed
            if (groupTiming.timeFilter) {
                const childStartProgress = childOffsetMs / groupTiming.durationMs
                const childEndProgress = firstChildOffsetMs / groupTiming.durationMs
                const reversePlayback = groupTiming.reversePlayback
                const parentStartProgress = timingProgress(groupTiming, reversePlayback ? 1 : 0)
                if (parentStartProgress >= childStartProgress && parentStartProgress <= childEndProgress) {
                    // Either playback direction can begin inside the child's active interval.
                    wallStartMs = 0
                } else {
                    // Enter from the interval boundary on the side the parent clock starts on.
                    // tmFilter still describes the authored forward curve when cTn@spd is negative;
                    // choose the first crossing in actual wall-clock order in either direction.
                    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
                    const entersFromAbove = parentStartProgress > childEndProgress
                    const threshold = entersFromAbove ? childEndProgress : childStartProgress
                    const reverseDirection = entersFromAbove ? 'down' : 'up'
                    const parentProgress = inverseTimingProgress(groupTiming, threshold, reversePlayback, reverseDirection)
                    const wallProgress = reversePlayback ? 1 - parentProgress : parentProgress
                    wallStartMs = Number.isFinite(wallProgress) ? wallProgress * scale : Infinity
                }
            } else {
                const wallProgress = inverseTimingProgress(groupTiming, firstChildOffsetMs / groupTiming.durationMs)
                wallStartMs = Number.isFinite(wallProgress)
                    ? (groupTiming.reversePlayback ? 1 - wallProgress : wallProgress) * scale
                    : Infinity
            }
        }
        return {
            ...action,
            delayMs: delayMs + wallStartMs,
            timingWarp: {
                acceleration: groupTiming.acceleration,
                deceleration: groupTiming.deceleration,
                timeFilter: groupTiming.timeFilter,
                groupOffsetMs: delayMs,
                groupDurationMs: groupTiming.durationMs,
                groupSpeed: groupTiming.playbackSpeed,
                groupReversePlayback: groupTiming.reversePlayback,
                groupAutoReverse: groupTiming.autoReverse,
                groupPlaybackDurationMs: groupTiming.playbackDurationMs,
                groupPerIteration: appliesPerIteration,
                childOffsetMs,
                groupEndDelayMs: delayMs + groupTiming.playbackDurationMs + (appliesPerIteration ? iterationSpan(action) : 0)
            }
        }
    })
    const offsetStartDelay = (conditions: Element[]) => {
        if (!conditions.length) return 0
        const delays: number[] = []
        for (const condition of conditions) {
            if (Array.from(condition.attributes).some(attribute => attribute.localName !== 'delay') || condition.children.length) return undefined
            const delay = condition.getAttribute('delay')
            if (isIndefiniteTimelineTime(delay) || (delay !== null && !isFiniteTimelineTime(delay))) return undefined
            delays.push(timelineTime(delay))
        }
        // Every condition in stCondLst must be met; offset-only conditions activate at the latest delay.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.startconditionlist?view=openxml-3.0.1
        return Math.max(...delays)
    }
    const startDelay = (timeNode: Element | undefined) => {
        const conditions = children(child(timeNode, 'stCondLst'), 'cond')
        // A start-condition list must be satisfied before its node can activate. An indefinite
        // delay can never be satisfied; do not let timelineTime() turn it into its zero fallback.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.startconditionlist?view=openxml-3.0.1
        if (conditions.some(condition => isIndefiniteTimelineTime(condition.getAttribute('delay')))) return Infinity
        const delay = offsetStartDelay(conditions)
        if (delay !== undefined) return delay
        warnings.add('An animation with event-based or invalid start conditions could not be activated.')
        return Infinity
    }
    const parseTextStyleValue = (property: TextStyleAnimationProperty, raw: string | null): TextStyleAnimationValue | undefined => {
        const rawValue = raw?.trim()
        if (property === 'fontFamily') {
            // PowerPoint permits arbitrary p:anim string values; retain punctuation in legitimate font names.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
            return rawValue && rawValue.length <= 128 && !/[\u0000-\u001f\u007f;]/u.test(rawValue) ? rawValue : undefined
        }
        const value = rawValue?.toLowerCase()
        if (property === 'fontWeight') return value === 'none' || value === 'normal' || value === 'bold' ? value : undefined
        if (property === 'fontStyle') return value === 'italic' ? true : value === 'normal' || value === 'none' ? false : undefined
        if (property === 'underline' || property === 'lineThrough' || property === 'outline') return value === 'true' ? true : value === 'false' ? false : undefined
        if (property === 'textShadow') return value === 'none' || value === 'normal' || value === 'auto' ? value : undefined
        if (property === 'emboss') return value === 'none' || value === 'normal' || value === 'emboss' ? value : undefined
        return value === 'none' || value === 'normal' || value === 'sub' || value === 'super' ? value : undefined
    }
    const hasCompleteTextOutline = (element: PptxElement) => {
        const runs = element.paragraphs.flatMap(paragraph => paragraph.runs).filter(run => run.text.length > 0)
        return runs.length > 0 && runs.every(run => run.outlineWidth !== undefined && run.outlineWidth > 0 && run.outlineColor !== undefined && run.outlineColor !== 'transparent')
    }
    const hasCompleteTextShadow = (element: PptxElement) => {
        const runs = element.paragraphs.flatMap(paragraph => paragraph.runs).filter(run => run.text.length > 0)
        return runs.length > 0 && runs.every(run => Boolean(run.shadow || run.shadows?.length))
    }
    const handledEffects = new Set<Element>()
    const handledScales = new Set<Element>()
    const handledRotations = new Set<Element>()
    const handledColors = new Set<Element>()
    const handledTextAnimations = new Set<Element>()
    const handledMotions = new Set<Element>()
    const handledMediaActions = new Set<Element>()
    // OOXML also defines withGroup/afterGroup as group-level relatives of withEffect/afterEffect.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timenodevalues?view=openxml-3.0.1
    const supportedGroupTypes = new Set(['clickEffect', 'clickPar', 'withEffect', 'afterEffect', 'withGroup', 'afterGroup'])
    const startsOnClick = (type: string | null | undefined) => type === 'clickEffect' || type === 'clickPar'
    const followsPreviousGroup = (type: string | null | undefined) => type === 'afterEffect' || type === 'afterGroup'
    // presetClass is a p:cTn attribute; discover nested preset nodes independently of a fixed timing-tree path.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode.presetclass?view=openxml-3.0.1
    const hasPresetClass = (timeNode: Element) => Boolean(trimXmlWhitespace(timeNode.getAttribute('presetClass')))
    const groupType = (timeNode: Element) => timeNode.getAttribute('nodeType') || (hasPresetClass(timeNode) ? 'clickEffect' : '')
    const isClickBuildWait = (timeNode: Element, type: string | null | undefined) => {
        const conditions = children(child(timeNode, 'stCondLst'), 'cond')
        return startsOnClick(type)
            && conditions.length === 1
            && conditions[0]?.getAttribute('delay') === 'indefinite'
            && Array.from(conditions[0].attributes).every(attribute => attribute.localName === 'delay')
            && conditions[0].children.length === 0
    }
    const sceneTree = firstDescendant(slideRoot, 'spTree')
    const elementsById = new Map<string, PptxElement>()
    for (const element of elements) if (!elementsById.has(element.id)) elementsById.set(element.id, element)
    const renderedElementIds = new Set(elementsById.keys())
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
    // Keep the owner node's own end separate from each action's earliest inherited parent end.
    const parsedTimeNodeEndOffsets = new Map<Element, number>()
    const parseGroup = (group: Element, delayMs: number, owner = child(group, 'cTn'), sequenceKey = 'main', sequenceOrder = 0, mediaWaitForEndKeys: string[] = [], mediaWaitDelayMs = 0, subTimeEntry?: Element, eventTimeNodeSources?: Map<string, {startMs: number; endMs?: number}>): PptxAnimation[] => {
        const resolvedStartDelay = (timeNode: Element | undefined) => {
            const conditions = children(child(timeNode, 'stCondLst'), 'cond')
            const offsetDelayMs = offsetStartDelay(conditions)
            if (offsetDelayMs !== undefined) return offsetDelayMs
            if (timeNode && eventTimeNodeSources) {
                const eventStartMs = subTimeEventStartTime(timeNode, delayMs, eventTimeNodeSources)
                if (eventStartMs !== undefined) return eventStartMs - delayMs
            }
            return startDelay(timeNode)
        }

        // Retain separate time nodes when several behaviors animate the same property.
        const actions = new Map<string, PptxAnimation>()
        // Keep valid timing when an unsupported visual behavior is skipped; a p:seq child starts after the prior child finishes.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.sequencetimenode?view=openxml-3.0.1
        const preserveTimingOnly = (
            targetId: string,
            timeNode: Element | undefined,
            timing: ReturnType<typeof animationTiming>
        ) => {
            if (!timing) return
            actions.set(`${targetId}:timing:${timeNode?.getAttribute('id') || actions.size}`, {
                targetId,
                sourceTimeNodeId: timeNode?.getAttribute('id') || undefined,
                effect: 'timingOnly',
                direction: 'in',
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + resolvedStartDelay(timeNode),
                ...timing
            })
        }
        const solidFillTypeTargets = new Set<string>()
        const mediaTimeNodeIds = new Map<string, string>()
        let groupTiming: GroupTiming | undefined
        const hasIteration = Boolean(child(owner, 'iterate'))
        const canApplyGroupTiming = (!subTimeEntry || supportsParallelSubTimeTiming(owner)) && Boolean(owner)
            && (['accel', 'decel', 'tmFilter', 'spd', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => owner!.hasAttribute(attribute))
                || Boolean(child(owner, 'endCondLst')))
        // Parent accel/decel/spd/autoRev and repeats shape the group clock; child offsets stay local to one cycle.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.childtimenodelist?view=openxml-2.19.0
        const iterate = subTimeEntry ? undefined : child(owner, 'iterate')
        // CT_TLIterateData defaults an omitted type to `el` (iterate over elements).
        // https://ecma-international.org/publications-and-standards/standards/ecma-376/
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.iterate?view=openxml-3.0.1
        const iterateType = () => iterate?.hasAttribute('type') ? iterate.getAttribute('type') || '' : 'el'
        const belongsToOwner = (node: Element) => {
            if (subTimeEntry ? !subTimeEntry.contains(node) : isInSubTimeNodeList(node)) return false
            let parent = node.parentElement
            while (parent && parent !== group.parentElement) {
                if (parent.localName === 'cTn' && ((subTimeEntry && parent === owner) || supportedGroupTypes.has(parent.getAttribute('nodeType') || '') || hasPresetClass(parent))) return parent === owner
                parent = parent.parentElement
            }
            return false
        }
        const ownedDescendants = (name: string) => {
            const nodes = descendants(group, name).filter(belongsToOwner)
            for (const node of nodes) {
                const accumulate = child(node, 'cBhvr')?.getAttribute('accumulate')
                if (accumulate === null || accumulate === undefined || accumulate === 'none') continue
                if (accumulate !== 'always') warnings.add('An animation behavior has an invalid cBhvr@accumulate value; it is ignored.')
                else {
                    const accumulatedProperties = new Set(['ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ScaleX', 'ScaleY', 'style.rotation', 'style.opacity', 'fill.opacity', 'stroke.opacity', 'shadow.opacity', 'shadow.offset.x', 'shadow.offset.y'])
                    const attributes = descendants(node, 'attrName').map(attribute => attribute.textContent?.trim() || '')
                    const supportedGenericAccumulation = name === 'anim' && attributes.length > 0 && attributes.every(attribute => accumulatedProperties.has(attribute))
                    if (!['animRot', 'animScale', 'animMotion'].includes(name) && !supportedGenericAccumulation) {
                        warnings.add(`The ${name} animation does not support repeated-value accumulation for its property.`)
                    }
                }
            }
            return nodes
        }
        // p:charRg uses an end-exclusive range; PowerPoint requires end > st and end <= text length.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d0b4598c-f4e1-47ff-91b0-8fb5abb708f3
        // p:txEl/p:pRg is the OOXML paragraph index range:
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.textelement.paragraphindexrange
        const target = (node: Element) => {
            const shape = firstDescendant(node, 'spTgt')
            const textElement = firstDescendant(shape, 'txEl')
            const paragraphNode = firstDescendant(textElement, 'pRg')
            const characterNode = firstDescendant(textElement, 'charRg')
            const diagramId = shape?.getAttribute('spid') || ''
            const subShapeId = child(shape, 'subSp')?.getAttribute('spid')
            const targetId = subShapeId ? smartArtAnimationTargets.get(`${diagramId}:${subShapeId}`) || '' : diagramId
            if (subShapeId && !targetId) warnings.add('A diagram subshape animation target could not be matched to a saved SmartArt shape and was skipped.')
            const start = numberAttr(characterNode, 'st')
            const end = numberAttr(characterNode, 'end')
            const entireText = start === 0xffff_ffff && end === 0xffff_ffff
            return {
                targetId,
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
            const intervalText = trimXmlWhitespace(rawInterval)
            const indefiniteAbsoluteInterval = intervalNode?.localName === 'tmAbs' && isIndefiniteTimelineTime(intervalText)
            // Office reads ST_PositivePercentage as thousandths of a percent or a percent with a % suffix.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c1f1feac-e34c-48d5-b2e2-f67bb67113e7
            const percentageInterval = intervalNode?.localName === 'tmPct' && intervalText && /^(?:\d+\.?\d*|\.\d+)%$/.test(intervalText)
                ? Number(intervalText.slice(0, -1)) * 1_000
                : NaN
            const intervalValue = indefiniteAbsoluteInterval ? Infinity : Number.isFinite(percentageInterval)
                ? percentageInterval
                : intervalText && isFiniteTimelineTime(intervalText) ? Number(intervalText) : NaN
            // tmAbs is an absolute wall-clock interval. The iterated child’s playback speed
            // changes its duration, but must not scale the gap between iterations.
            // TimeAbsolute.val uses ST_TLTime, which also permits `indefinite`.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timeabsolute.val?view=openxml-3.0.1
            const intervalMs = indefiniteAbsoluteInterval ? Infinity
                : intervalNode?.localName === 'tmAbs' ? (intervalValue <= MAX_TL_TIME_MS ? intervalValue : NaN)
                    : intervalNode?.localName === 'tmPct' ? (iterationDurationMs * intervalValue) / 100_000 : NaN
            const backwardsValue = iterate.getAttribute('backwards')
            // PowerPoint clarifies that `backwards` chooses iteration from the target's end; it does not reverse each child effect.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/b0d3dac1-75dd-4212-8c7d-af9f432b5c24
            const backwards = backwardsValue === '1' || backwardsValue === 'true'
            if (intervalNodes.length !== 1 || (!indefiniteAbsoluteInterval && (!Number.isFinite(intervalValue) || intervalValue < 0 || intervalValue > MAX_TL_TIME_MS || !Number.isFinite(intervalMs) || intervalMs < 0))) {
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
            const type = iterateType()
            if (type === 'el') return undefined
            const interval = iterationTiming(action)
            const element = elementsById.get(action.targetId)
            if (!['lt', 'wd'].includes(type || '') || !interval || !element?.paragraphs.some(paragraph => paragraph.runs.length)) {
                if (!['lt', 'wd'].includes(type || '')) warnings.add('A text iteration with an unsupported type was rendered as a single animation.')
                else if (!element?.paragraphs.some(paragraph => paragraph.runs.length)) warnings.add('A text iteration target has no text and was rendered as a single animation.')
                return undefined
            }
            const {intervalMs, backwards} = interval
            if (action.effect === 'color' && action.colorProperty === 'fillcolor' && element.kind === 'shape') {
                warnings.add('Text iteration does not stagger shape fill color; the fill animation is rendered once across the shape.')
                return undefined
            }
            // p:iterate applies child animation behavior to each letter or word.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.iterate.type?view=openxml-3.0.1
            if (!['appear', 'fade', 'flash', 'blinds', 'checker', 'dissolve', 'wheel', 'wedge', 'randomBars', 'shape', 'barn', 'wipe', 'slide', 'strips', 'scale', 'zoom', 'spiral', 'rotation', 'motion', 'fontSize', 'fontWeight', 'textStyle', 'opacity', 'shadowOffset', 'visibility'].includes(action.effect) && !(action.effect === 'color' && ['style.color', 'shadow.color', 'stroke.color'].includes(action.colorProperty || ''))) {
                warnings.add(action.effect === 'color'
                    ? `Text iteration for ${action.colorProperty || 'color'} animation is not rendered yet.`
                    : `Text iteration for ${action.effect} animation is not rendered yet.`)
                return undefined
            }
            if (!Intl.Segmenter) {
                warnings.add('Text iteration is unavailable because Intl.Segmenter is not supported by this browser.')
                return undefined
            }
            const granularity = type === 'lt' ? 'grapheme' : 'word'
            const segmenters = new Map<string, Intl.Segmenter>()
            const segmenterFor = (paragraph: PptxParagraph) => {
                const locale = type === 'wd' ? textIterationLocale(paragraph.runs) : undefined
                const key = locale || ''
                let segmenter = segmenters.get(key)
                if (!segmenter) {
                    segmenter = createTextIterationSegmenter(locale, granularity)
                    segmenters.set(key, segmenter)
                }
                return segmenter
            }
            const ranges: Array<{start: number; end: number}> = []
            let overflow = false
            let paragraphOffset = 0
            for (const [paragraphIndex, paragraph] of element.paragraphs.entries()) {
                const text = paragraph.runs.map(run => run.text).join('')
                const paragraphSelected = !action.paragraphRange || (paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
                if (paragraphSelected) {
                    for (const range of textIterationRanges(type, segmenterFor(paragraph).segment(text))) {
                        // Whitespace keeps its source offset but does not consume an iteration slot.
                        // Word punctuation joins its adjacent word so it does not appear ahead of that word.
                        const start = paragraphOffset + range.start
                        const end = paragraphOffset + range.end
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
            if (!iterate || iterateType() !== 'el') return undefined
            const targetIds = groupChildren(action.targetId) || (renderedElementIds.has(action.targetId) ? [action.targetId] : [])
            if (!targetIds.length) {
                warnings.add('A shape iteration target has no rendered shapes and was skipped.')
                return []
            }
            if (targetIds.length === 1) return [{...action, targetId: targetIds[0]!}]
            const interval = iterationTiming(action, true)
            return targetIds.map((targetId, index) => {
                const iterationIndex = interval?.backwards ? targetIds.length - index - 1 : index
                // `indefinite` leaves the first iterated shape at the group start; Infinity * 0 would yield NaN.
                const staggerMs = interval && iterationIndex > 0 ? interval.intervalMs * iterationIndex : 0
                return {
                    ...action,
                    targetId,
                    delayMs: action.delayMs + staggerMs
                }
            })
        }
        const addMediaAction = (node: Element, targetId: string, mediaCommand: NonNullable<PptxAnimation['mediaCommand']>, options: Pick<PptxAnimation, 'mediaStartSeconds' | 'mediaSourceUrl' | 'mediaVolume' | 'mediaMuted' | 'mediaDurationMs' | 'mediaSlideCount' | 'mediaWaitForEnd' | 'mediaIsNarration' | 'mediaIsInstantCommand'> = {}) => {
            const element = elementsById.get(targetId)
            if (mediaCommand !== 'stopAudio' && !options.mediaSourceUrl && (!element || !['audio', 'video'].includes(element.kind))) {
                warnings.add('A media timeline action without an embedded audio or video target was skipped.')
                handledMediaActions.add(node)
                return
            }
            const timeNode = firstDescendant(node, 'cTn')
            const rawPlaybackSpeed = timeNode?.getAttribute('spd')
            // ISO/IEC 29500 defines cTn@spd as a timing-rate percentage; 200% halves the duration.
            // Browser media engines do not reliably support reverse playback for both audio and video.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
            const parsedPlaybackSpeed = parseAnimationSpeed(rawPlaybackSpeed)
            if (parsedPlaybackSpeed === undefined) warnings.add('A media timeline with an invalid speed uses normal playback speed.')
            if ((parsedPlaybackSpeed ?? 1) < 0) warnings.add('A media timeline with negative playback speed uses normal playback speed because reverse audio/video playback is not supported.')
            const mediaPlaybackRate = parsedPlaybackSpeed === undefined || parsedPlaybackSpeed < 0 ? 1 : parsedPlaybackSpeed
            const conditions = children(child(timeNode, 'stCondLst'), 'cond')
            const condition = conditions[0]
            const delay = condition?.getAttribute('delay')
            const parentBeginCondition = Boolean(
                subTimeEntry?.contains(node) && owner?.getAttribute('id') && condition &&
                ['begin', 'onBegin'].includes(condition.getAttribute('evt') || '') &&
                (delay === null || isFiniteTimelineTime(delay)) &&
                Array.from(condition.attributes).every(attribute => ['evt', 'delay'].includes(attribute.localName)) &&
                condition.children.length === 1 &&
                firstDescendant(condition, 'tn')?.getAttribute('val') === owner.getAttribute('id')
            )
            const simpleDelayCondition = !condition || (
                Array.from(condition.attributes).every(attribute => attribute.localName === 'delay') &&
                condition.children.length === 0 &&
                !isIndefiniteTimelineTime(delay) &&
                (delay === null || isFiniteTimelineTime(delay))
            )
            if (conditions.length > 1 || (condition && !simpleDelayCondition && !parentBeginCondition)) {
                warnings.add('A media timeline action with event-based or invalid start conditions was skipped.')
                handledMediaActions.add(node)
                return
            }
            const timeNodeId = timeNode?.getAttribute('id') || undefined
            const key = `media:${timeNodeId || actions.size}`
            if (timeNodeId) mediaTimeNodeIds.set(key, timeNodeId)
            actions.set(key, {
                targetId,
                effect: 'media',
                direction: 'in',
                mediaCommand,
                ...options,
                mediaPlaybackRate,
                mediaDurationMs: options.mediaDurationMs === undefined
                    ? undefined
                    : options.mediaDurationMs === 0
                        ? 0
                        : mediaPlaybackRate === 0
                            ? Infinity
                            : options.mediaDurationMs / mediaPlaybackRate,
                mediaWaitForEnd: options.mediaWaitForEnd || (mediaCommand === 'play' && mediaPlaybackRate === 0 && options.mediaDurationMs !== 0),
                durationMs: 0,
                delayMs: delayMs + timelineTime(delay)
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
            const soundTarget = child(child(media, 'tgtEl'), 'sndTgt')
            const mediaSourceUrl = soundTarget ? soundTargetUrls.get(mediaNode) : undefined
            const isSoundTarget = Boolean(soundTarget)
            const targetId = isSoundTarget ? soundTargetIds.get(mediaNode) : animationTarget.targetId
            if (!targetId || (!isSoundTarget && (animationTarget.hasUnsupportedTextRange || animationTarget.paragraphRange || animationTarget.characterRange)) || (isSoundTarget && !mediaSourceUrl)) {
                warnings.add('A timed media playback action with an unsupported target was skipped.')
                handledMediaActions.add(mediaNode)
                continue
            }
            const volumeValue = media.getAttribute('vol')
            // CT_TLCommonMediaNodeData defaults vol to 50%; browser media elements default to 100%.
            // https://ecma-international.org/publications-and-standards/standards/ecma-376/
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonmedianode?view=openxml-3.0.1
            const volume = volumeValue === null ? 0.5 : Number(volumeValue) / 100_000
            const muteValue = media.getAttribute('mute')?.toLowerCase()
            const muted = muteValue === undefined ? undefined : ['1', 'true'].includes(muteValue)
            const slideCountValue = media.getAttribute('numSld')
            const parsedSlideCount = slideCountValue === null ? undefined : Number(slideCountValue)
            const mediaSlideCount = parsedSlideCount !== undefined && Number.isInteger(parsedSlideCount) && parsedSlideCount > 0 && parsedSlideCount <= 0xffff_ffff ? parsedSlideCount : undefined
            if (volumeValue !== null && (!Number.isFinite(volume) || volume! < 0 || volume! > 1)) warnings.add('An invalid media timeline volume was ignored.')
            if (muteValue !== undefined && !['0', '1', 'false', 'true'].includes(muteValue)) warnings.add('An invalid media timeline mute value was ignored.')
            if (slideCountValue !== null && mediaSlideCount === undefined) warnings.add('An invalid media timeline numSld value was ignored.')
            const mediaTargetElement = elementsById.get(animationTarget.targetId)
            const narrationValue = mediaNode.getAttribute('isNarration')?.trim().toLowerCase()
            const requestedNarration = narrationValue === '1' || narrationValue === 'true'
            if (narrationValue !== undefined && !['0', '1', 'false', 'true'].includes(narrationValue)) warnings.add('An invalid p:audio isNarration value was ignored.')
            const mediaIsNarration = requestedNarration && (isSoundTarget || mediaTargetElement?.kind === 'audio') ? true : undefined
            if (requestedNarration && mediaIsNarration === undefined) warnings.add('A p:audio isNarration flag without an audio shape was ignored.')
            if (mediaSlideCount && mediaSlideCount > 1 && mediaTargetElement?.kind === 'video' && (mediaTargetElement.mediaTrim || mediaTargetElement.mediaFade)) warnings.add('Cross-slide video playback ignores video trim and fades, as documented for PowerPoint 2010/2013.')
            const mediaTimeNode = firstDescendant(media, 'cTn')
            const mediaDurationValue = mediaTimeNode?.getAttribute('dur')
            const endConditions = children(child(mediaTimeNode, 'endCondLst'), 'cond')
            const hasSupportedEndCondition =
                (isSoundTarget || mediaTargetElement?.kind === 'audio') &&
                endConditions.every(condition => {
                    const delay = condition.getAttribute('delay')
                    return condition.getAttribute('evt') === 'onStopAudio' && Boolean(firstDescendant(condition, 'sldTgt')) && (delay === null || (isFiniteTimelineTime(delay) && timelineTime(delay, Number.NaN) === 0))
                })
            const indefiniteDuration = isIndefiniteTimelineTime(mediaDurationValue)
            // WPS/PowerPoint sound-effect cMediaNodes commonly omit dur and end with the embedded media.
            const implicitMediaDuration = mediaDurationValue === null || mediaDurationValue === undefined
            const parsedMediaDuration = implicitMediaDuration ? undefined : isFiniteTimelineTime(mediaDurationValue) ? Number(mediaDurationValue.trim()) : undefined
            const mediaDurationMs = parsedMediaDuration !== undefined && Number.isFinite(parsedMediaDuration) && parsedMediaDuration <= MAX_TL_TIME_MS ? parsedMediaDuration : undefined
            if (parsedMediaDuration !== undefined && mediaDurationMs === undefined) warnings.add('A media timeline duration beyond the supported ST_TLTime range was skipped.')
            if (mediaDurationValue !== null && mediaDurationValue !== undefined && !indefiniteDuration && parsedMediaDuration === undefined) warnings.add('An invalid media timeline duration was skipped.')
            if (endConditions.length && !hasSupportedEndCondition) warnings.add('A media timeline end condition other than slide-targeted onStopAudio is not synchronized yet.')
            addMediaAction(mediaNode, targetId, 'play', {
                mediaSourceUrl,
                mediaVolume: volume !== undefined && volume >= 0 && volume <= 1 ? volume : undefined,
                mediaMuted: muteValue === undefined || ['0', '1', 'false', 'true'].includes(muteValue) ? muted : undefined,
                mediaDurationMs,
                mediaSlideCount,
                mediaIsNarration,
                mediaWaitForEnd: (indefiniteDuration || implicitMediaDuration) && (!endConditions.length || hasSupportedEndCondition)
            })
        }
        for (const command of ownedDescendants('cmd').filter(node => node.namespaceURI === PRESENTATION_NS)) {
            const animationTarget = target(command)
            const commandType = command.getAttribute('type')
            const source = command.getAttribute('cmd')?.trim() || ''
            const playFrom = /^PlayFrom\(\s*(\d+(?:\.\d+)?)\s*\)$/i.exec(source)
            const normalizedCommand = source.toLowerCase()
            // MS-OE376 §4.6.28: type="evt" cmd="onstopaudio" stops all audio.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b2f15f11-5b40-4f67-ae95-5a510c3bea06
            if (commandType === 'evt' && normalizedCommand === 'onstopaudio') {
                if (!firstDescendant(command, 'sldTgt') || animationTarget.hasUnsupportedTextRange || animationTarget.paragraphRange || animationTarget.characterRange) {
                    warnings.add('A stop-all-audio event without a slide target was skipped.')
                    handledMediaActions.add(command)
                    continue
                }
                addMediaAction(command, '', 'stopAudio', {mediaIsInstantCommand: true})
                continue
            }
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
            // p:cmd represents a non-durational command, so its begin instant remains meaningful at the parent end boundary.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.command?view=openxml-3.0.1
            addMediaAction(command, animationTarget.targetId, mediaCommand, {mediaStartSeconds: startSeconds, mediaIsInstantCommand: true})
        }
        for (const [effectIndex, effect] of ownedDescendants('animEffect').entries()) {
            const animationTarget = target(effect)
            const textElement = elementsById.get(animationTarget.targetId)
            const effectTimeNode = firstDescendant(effect, 'cTn')
            const preserveEffectTiming = () => preserveTimingOnly(animationTarget.targetId, effectTimeNode, animationTiming(effectTimeNode, 'Object effect'))
            // ECMA-376 CT_TLAnimateEffectBehavior declares transition default="in".
            const transition = trimXmlWhitespace(effect.getAttribute('transition')) ?? 'in'
            // PowerPoint interprets ST_TLAnimateEffectTransition="none" as "in".
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b4fa69af-99d9-4423-a6f9-699b95b13238
            if (!['in', 'out', 'none'].includes(transition)) {
                warnings.add(`The object-effect transition "${transition || '(empty)'}" is unsupported; its timing was preserved without rendering the effect.`)
                preserveEffectTiming()
                handledEffects.add(effect)
                continue
            }
            const normalizedTransition = transition === 'none' ? 'in' : transition
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
            const supportedFilter = ['fade', 'cut'].includes(filter || '') || Boolean(wipe || blinds || checker || randomBar || strips || barn || slide || dissolve || shapeFilter || wheel || wedge)
            const presetIdText = trimXmlWhitespace(owner?.getAttribute('presetID')) || ''
            const presetSubtypeSpecified = owner?.hasAttribute('presetSubtype') ?? false
            const presetSubtypeText = trimXmlWhitespace(owner?.getAttribute('presetSubtype')) ?? ''
            const presetClass = trimXmlWhitespace(owner?.getAttribute('presetClass'))
            const presetId = /^[+-]?\d+$/.test(presetIdText) ? Number(presetIdText) : undefined
            const presetSubtype = /^[+-]?\d+$/.test(presetSubtypeText) ? Number(presetSubtypeText) : undefined
            const matchingPresetClass = (normalizedTransition === 'in' && presetClass === 'entr') || (normalizedTransition === 'out' && presetClass === 'exit')
            // Resolve known preset metadata only when animEffect has no supported filter.
            const presetAppear = !supportedFilter && matchingPresetClass && presetId === 1
            // MS-PPT assigns Flash Once effect ID 11 and says its direction is ignored.
            // This browser approximation emits one on/off blink before settling at the authored endpoint.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
            const presetFlash = !supportedFilter && matchingPresetClass && presetId === 11
            if (presetFlash) warnings.add('A Flash Once preset is approximated as one discrete opacity blink.')
            // MS-PPT assigns Spiral effect ID 15 and ignores its direction value.
            // The browser approximation composes one centered turn with a scale reveal/conceal.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
            const presetSpiral = !supportedFilter && matchingPresetClass && presetId === 15
            // Whole-shape transforms are applied on top of the authored shape rotation.
            // Text ranges and per-word/per-letter transforms run inside the already-rotated shape.
            const presetSpiralBaseRotation = animationTarget.paragraphRange || animationTarget.characterRange || ['lt', 'wd'].includes(iterate?.getAttribute('type') || '')
                ? 0
                : textElement?.rotation || 0
            if (presetSpiral) warnings.add('A Spiral preset is approximated with a centered spin-and-scale.')
            // PowerPoint lists Fade, Wheel, and Wipe as effects 10, 21, and 22; subtype normalization follows the supplied parser reference.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            const presetFade = !supportedFilter && matchingPresetClass && presetId === 10
            const presetWheel = !supportedFilter && matchingPresetClass && presetId === 21
            const presetWheelSpokes = presetWheel
                ? WHEEL_PRESET_SPOKES.get(presetSubtypeSpecified ? presetSubtype ?? Number.NaN : 1) ?? 1
                : undefined
            if (presetWheel && presetSubtypeSpecified && !WHEEL_PRESET_SPOKES.has(presetSubtype ?? Number.NaN)) {
                warnings.add(`The Wheel preset subtype "${presetSubtypeText || '(invalid)'}" is unsupported; it was approximated as Wheel(1).`)
            }
            const presetWipe = !supportedFilter && matchingPresetClass && presetId === 22
            const presetWipeDirection = presetWipe
                ? WIPE_PRESET_DIRECTIONS.get(presetSubtypeSpecified ? presetSubtype ?? Number.NaN : 4) ?? 'down'
                : undefined
            if (presetWipe && presetSubtypeSpecified && !WIPE_PRESET_DIRECTIONS.has(presetSubtype ?? Number.NaN)) {
                warnings.add(`The Wipe preset subtype "${presetSubtypeText || '(invalid)'}" is unsupported; it was approximated as a downward wipe.`)
            }
            // Microsoft MsoAnimEffect IDs 9 and 20 are Dissolve and Wedge; MS-OE376 lists their matching object filters.
            const presetDissolve = !supportedFilter && matchingPresetClass && presetId === 9
            const presetWedge = !supportedFilter && matchingPresetClass && presetId === 20
            // The supplied Fly preset matrix uses bottom (bit 4) when presetSubtype is omitted.
            const flyPresetSubtype = presetSubtype ?? (owner?.hasAttribute('presetSubtype') ? Number.NaN : 4)
            // Microsoft defines presetSubtype as a direction bitflag; the supplied parser reference maps Fly's eight directional values.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode?view=openxml-3.0.1
            const presetFlyDirection = !supportedFilter && matchingPresetClass && presetId === 2 ? FLY_PRESET_DIRECTIONS.get(flyPresetSubtype) : undefined
            // Microsoft assigns Crawl presetID=7 and uses the same eight direction flags as Fly.
            // The browser player approximates Crawl with the existing directional slide motion.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
            const crawlPresetSubtype = presetSubtype ?? (presetSubtypeSpecified ? Number.NaN : 4)
            const presetCrawlDirection = !supportedFilter && matchingPresetClass && presetId === 7 ? FLY_PRESET_DIRECTIONS.get(crawlPresetSubtype) : undefined
            if (presetCrawlDirection) warnings.add('A Crawl preset is approximated as directional slide motion.')
            // MS-PPT defines Stretch as preset 17 with direction subtypes Right/Left/Top/Bottom/Across.
            // Scale around the opposite edge to preserve the direction of growth.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
            const stretchPresetSubtype = presetSubtype ?? (presetSubtypeSpecified ? Number.NaN : 10)
            const presetStretch = !supportedFilter && matchingPresetClass && presetId === 17
                ? STRETCH_PRESET_DIRECTIONS.get(stretchPresetSubtype)
                : undefined
            if (!presetStretch && !supportedFilter && matchingPresetClass && presetId === 17) {
                warnings.add('A Stretch preset with an unsupported direction was approximated as a fade.')
            }
            // MS-PPT defines Swivel as preset 19 with horizontal (10) or vertical (5) direction.
            // The browser player maps those axes to a CSS 3D half-turn.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
            const swivelPresetSubtype = presetSubtype ?? (presetSubtypeSpecified ? Number.NaN : 10)
            const presetSwivelAxis = !supportedFilter && matchingPresetClass && presetId === 19
                ? SWIVEL_PRESET_AXES.get(swivelPresetSubtype)
                : undefined
            if (presetSwivelAxis) warnings.add('A Swivel preset is approximated with a CSS 3D rotation.')
            if (!presetSwivelAxis && !supportedFilter && matchingPresetClass && presetId === 19) {
                warnings.add('A Swivel preset with an unsupported direction was approximated as a fade.')
            }
            const peekPresetSubtype = presetSubtype ?? (presetSubtypeSpecified ? Number.NaN : 4)
            // Peek In/Out slides the object into or out of view from one of the four cardinals.
            // https://support.microsoft.com/en-au/powerpoint/animation-effects-available-in-powerpoint-for-the-web
            const presetPeekDirection = !supportedFilter && matchingPresetClass && presetId === 12 ? PEEK_PRESET_DIRECTIONS.get(peekPresetSubtype) : undefined
            const presetSplit = !supportedFilter && matchingPresetClass && presetId === 16 ? SPLIT_PRESET_DIRECTIONS.get(presetSubtype ?? 21) : undefined
            const presetSplitMotion = presetSplit && presetClass === 'exit'
                ? presetSplit.motion === 'in' ? 'out' : 'in'
                : presetSplit?.motion
            // MS-PPT defines Zoom subtype values 16/32, 528, 36, 288, and 272.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/35c58b56-5532-4d77-8c33-53a82efb5a89
            const presetZoomVariant = !supportedFilter && matchingPresetClass && presetId === 23
                ? !presetSubtypeSpecified
                    ? normalizedTransition === 'in' ? {direction: 'in' as const, from: 0, to: 1} : {direction: 'out' as const, from: 1, to: 0}
                    : presetSubtype === undefined ? undefined : ZOOM_PRESET_SUBTYPES.get(presetSubtype)
                : undefined
            const presetZoomDirection = presetZoomVariant?.direction
            const presetZoomScaleFrom = presetZoomVariant ? [presetZoomVariant.from, presetZoomVariant.from] as [number, number] : undefined
            const presetZoomScaleTo = presetZoomVariant ? [presetZoomVariant.to, presetZoomVariant.to] as [number, number] : undefined
            const presetZoomOrigin = presetZoomVariant?.origin && textElement && textElement.width > 0 && textElement.height > 0
                ? [
                    ((slideWidth / 2) - textElement.x) / textElement.width,
                    ((presetZoomVariant.origin === 'screenCenter' ? slideHeight / 2 : slideHeight) - textElement.y) / textElement.height
                ] as [number, number]
                : undefined
            if (presetZoomVariant?.approximate) warnings.add('A Zoom Slightly preset is approximated with a 0.8 scale endpoint.')
            if (presetZoomVariant?.origin) warnings.add('A screen-origin Zoom preset is approximated by scaling around its named slide point.')
            if (!presetZoomVariant && !supportedFilter && matchingPresetClass && presetId === 23) {
                warnings.add(`The Zoom preset subtype "${presetSubtypeText || '(missing)'}" is unsupported; it falls back to fade.`)
            }
            // Office-compatibility inference: cTn/@presetID is an OOXML integer and PowerPoint's MsoAnimEffect lists Rise Up as 34.
            // Keep that distinct from MS-PPT's separate TimeEffectID table for the legacy binary .ppt format.
            // Approximate the preset with the existing bottom-to-position slide motion.
            // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commontimenode.presetid?view=openxml-3.0.1
            const presetRiseUp = !supportedFilter && matchingPresetClass && presetId === 34
            if (presetRiseUp) warnings.add('A Rise Up preset is approximated as a vertical slide from the bottom.')
            const presetShape = !supportedFilter && matchingPresetClass && presetId !== undefined ? SHAPE_ANIMATION_PRESET_IDS.get(presetId) : undefined
            const presetShapeDirection = presetShape ? getShapePresetDirection(normalizedTransition as 'in' | 'out', presetSubtype) : undefined
            // The Office preset IDs for Blinds, Checkerboard, Random Bars, and Strips
            // match the existing filter renderers; an omitted subtype uses each effect's common default.
            const presetBlinds = !supportedFilter && matchingPresetClass && presetId === 3
            const presetChecker = !supportedFilter && matchingPresetClass && presetId === 5
            const presetRandomBar = !supportedFilter && matchingPresetClass && presetId === 14
            const presetStrips = !supportedFilter && matchingPresetClass && presetId === 18
            const presetBlindsOrientation = presetBlinds ? presetVariantOrDefault(presetSubtype, presetSubtypeText, presetSubtypeSpecified, PRESET_BLINDS_ORIENTATIONS, 'horizontal', 'Blinds', warnings) : undefined
            const presetCheckerOrientation = presetChecker ? presetVariantOrDefault(presetSubtype, presetSubtypeText, presetSubtypeSpecified, PRESET_CHECKER_ORIENTATIONS, 'horz', 'Checkerboard', warnings) : undefined
            const presetRandomBarOrientation = presetRandomBar ? presetVariantOrDefault(presetSubtype, presetSubtypeText, presetSubtypeSpecified, PRESET_RANDOM_BAR_ORIENTATIONS, 'horizontal', 'Random Bars', warnings) : undefined
            const presetStripsDirection = presetStrips ? presetVariantOrDefault(presetSubtype, presetSubtypeText, presetSubtypeSpecified, PRESET_STRIPS_DIRECTIONS, 'rd', 'Strips', warnings) : undefined
            const fadeFallback = ['in', 'out'].includes(normalizedTransition || '') && !supportedFilter && !presetFlyDirection && !presetCrawlDirection && !presetPeekDirection && !presetRiseUp && !presetStretch && !presetSwivelAxis && !presetFlash && !presetSpiral && !presetAppear && !presetFade && !presetWheel && !presetWipe && !presetSplit && !presetZoomDirection && !presetShape && !presetDissolve && !presetWedge && !presetBlinds && !presetChecker && !presetRandomBar && !presetStrips
            if (!animationTarget.targetId) {
                warnings.add('An object animation has no shape target and was skipped.')
                preserveEffectTiming()
                continue
            }
            if (animationTarget.hasUnsupportedTextRange) {
                warnings.add('This text range animation target is malformed or unsupported.')
                preserveEffectTiming()
                continue
            }
            if (!hasValidTextRange(animationTarget, textElement)) {
                warnings.add('A text-range animation exceeds the target text bounds and was skipped.')
                preserveEffectTiming()
                continue
            }
            // WPS Transparency exports an image filter with a fixed opacity property list.
            // Keep this as a property behavior so it does not become an entrance/exit fade.
            const imageProperties = filter === 'image' ? effect.getAttribute('prLst') || '' : ''
            const imageOpacityMatch = imageProperties.length <= 256
                ? /^\s*opacity\s*:\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*;?\s*$/i.exec(imageProperties)
                : null
            const imageOpacity = imageOpacityMatch ? Number(imageOpacityMatch[1]) : Number.NaN
            if (filter === 'image' && Number.isFinite(imageOpacity) && imageOpacity >= 0 && imageOpacity <= 1) {
                const behavior = child(effect, 'cBhvr')
                const timeNode = child(behavior, 'cTn') || firstDescendant(effect, 'cTn')
                const timing = animationTiming(timeNode, 'Image opacity')
                if (!timing) continue
                actions.set(`${animationTarget.targetId}:image-opacity:${timeNode?.getAttribute('id') || effectIndex}`, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'opacity',
                    opacityProperty: 'style.opacity',
                    opacityFrom: imageOpacity,
                    opacityTo: imageOpacity,
                    direction: 'in',
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledEffects.add(effect)
                continue
            }
            // MS-OE376 names the supported object filters; unknown or missing filters still honor in/out as a warned fade.
            if (['in', 'out'].includes(normalizedTransition || '')) {
                if (fadeFallback) warnings.add('An unsupported object-effect filter was approximated as a fade.')
                const paragraphWipe = Boolean(wipe && animationTarget.paragraphRange && !animationTarget.characterRange)
                const characterWipe = Boolean(wipe && animationTarget.characterRange && !animationTarget.paragraphRange)
                const iteratedText = ['lt', 'wd'].includes(iterate?.getAttribute('type') || '') && Boolean(textElement?.paragraphs.some(paragraph => paragraph.runs.length))
                const textRangeMask = Boolean(animationTarget.paragraphRange) !== Boolean(animationTarget.characterRange) || iteratedText
                const textRandomBars = Boolean(randomBar && textRangeMask)
                const textStrips = Boolean(strips && textRangeMask)
                const textBarn = Boolean(barn && textRangeMask)
                const textShapeMask = Boolean((dissolve || presetDissolve || shapeFilter || presetShape || wheel || wedge || presetWedge) && textRangeMask)
                // Microsoft documents pRg with checkerboard(across) as a valid paragraph-targeted animation.
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.paragraphindexrange?view=openxml-3.0.1
                const unsupportedTextMask = (randomBar && !textRandomBars) || (strips && !textStrips) || (barn && !textBarn) || ((dissolve || presetDissolve || shapeFilter || presetShape || wheel || wedge || presetWedge) && !textShapeMask) || (wipe && !paragraphWipe && !characterWipe)
                if (unsupportedTextMask && (animationTarget.paragraphRange || animationTarget.characterRange)) {
                    warnings.add('An unsupported text-range mask combination was skipped.')
                    preserveEffectTiming()
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
                const effectProgressNode = child(effect, 'progress')
                const effectProgressValues = effectProgressNode ? children(effectProgressNode) : []
                const effectProgressRaw = effectProgressValues.length === 1 && effectProgressValues[0]?.localName === 'fltVal'
                    ? effectProgressValues[0].getAttribute('val')?.trim()
                    : undefined
                const parsedEffectProgress = effectProgressRaw ? Number(effectProgressRaw) : Number.NaN
                const effectProgress = effectProgressNode
                    ? Number.isFinite(parsedEffectProgress) && parsedEffectProgress >= 0 && parsedEffectProgress <= 1
                        ? parsedEffectProgress
                        : undefined
                    // A cut filter changes visibility instantaneously; keep the authored duration
                    // for sequence scheduling while holding the visual effect at its end state.
                    : filter === 'cut' ? 1 : undefined
                if (effectProgressNode && effectProgress === undefined) {
                    warnings.add('An object-effect progress override must be a single float from 0 to 1; the timed progress was used instead.')
                }
                const key = `${animationTarget.targetId}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:effect:${effectIndex}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effectProgress,
                    effect: fadeFallback || presetFade ? 'fade' : presetFlash ? 'flash' : presetSpiral ? 'spiral' : presetStretch ? 'scale' : presetSwivelAxis ? 'rotation' : presetZoomDirection ? 'zoom' : wedge || presetWedge ? 'wedge' : wheel || presetWheel ? 'wheel' : slide || presetFlyDirection || presetCrawlDirection || presetPeekDirection || presetRiseUp ? 'slide' : strips || presetStrips ? 'strips' : barn || presetSplit ? 'barn' : dissolve || presetDissolve ? 'dissolve' : shapeFilter || presetShape ? 'shape' : randomBar || presetRandomBar ? 'randomBars' : checker || presetChecker ? 'checker' : blinds || presetBlinds ? 'blinds' : wipe || presetWipe ? 'wipe' : filter === 'fade' ? 'fade' : 'appear',
                    direction: presetZoomDirection || normalizedTransition as 'in' | 'out',
                    scaleFrom: presetSpiral
                        ? normalizedTransition === 'in' ? [0, 0] : [1, 1]
                        : presetStretch
                        ? normalizedTransition === 'in' ? [presetStretch.axis === 'y' ? 1 : 0, presetStretch.axis === 'x' ? 1 : 0] : [1, 1]
                        : presetZoomScaleFrom,
                    scaleTo: presetSpiral
                        ? normalizedTransition === 'in' ? [1, 1] : [0, 0]
                        : presetStretch
                        ? normalizedTransition === 'in' ? [1, 1] : [presetStretch.axis === 'y' ? 1 : 0, presetStretch.axis === 'x' ? 1 : 0]
                        : presetZoomScaleTo,
                    scaleOrigin: presetStretch?.origin || presetZoomOrigin,
                    rotationFrom: presetSpiral ? presetSpiralBaseRotation + (normalizedTransition === 'in' ? 360 : 0) : presetSwivelAxis ? normalizedTransition === 'in' ? 90 : 0 : undefined,
                    rotationTo: presetSpiral ? presetSpiralBaseRotation + (normalizedTransition === 'in' ? 0 : 360) : presetSwivelAxis ? normalizedTransition === 'in' ? 0 : 90 : undefined,
                    rotationAxis: presetSwivelAxis,
                    opacityFrom: presetFlash ? normalizedTransition === 'in' ? 0 : 1 : undefined,
                    opacityTo: presetFlash ? normalizedTransition === 'in' ? 1 : 0 : undefined,
                    opacityKeyframes: presetFlash
                        ? normalizedTransition === 'in'
                            ? [{offset: 0, value: 0}, {offset: 0.25, value: 1}, {offset: 0.5, value: 0}, {offset: 0.75, value: 1}, {offset: 1, value: 1}]
                            : [{offset: 0, value: 1}, {offset: 0.25, value: 0}, {offset: 0.5, value: 1}, {offset: 0.75, value: 0}, {offset: 1, value: 0}]
                        : undefined,
                    opacityKeyframeMode: presetFlash ? 'discrete' : undefined,
                    durationMs: presetAppear ? 0 : animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing,
                    wipeDirection: wipe ? wipe[1] as PptxAnimation['wipeDirection'] : presetWipeDirection,
                    blindsOrientation: blinds?.[1] as PptxAnimation['blindsOrientation'] ?? presetBlindsOrientation,
                    checkerOrientation: checker?.[1] === 'across' ? 'horz' : checker ? 'vert' : presetCheckerOrientation,
                    randomBarOrientation: randomBar?.[1] as PptxAnimation['randomBarOrientation'] ?? presetRandomBarOrientation,
                    stripsDirection: strips ? ({downLeft: 'ld', upLeft: 'lu', downRight: 'rd', upRight: 'ru'} as const)[strips[1] as 'downLeft' | 'upLeft' | 'downRight' | 'upRight'] : presetStripsDirection,
                    barnOrientation: barn ? barn[2]?.toLowerCase() as PptxAnimation['barnOrientation'] : presetSplit?.orientation,
                    barnMotion: barn ? barn[1] as PptxAnimation['barnMotion'] : presetSplitMotion,
                    shapeFilter: (shapeFilter?.[1] as PptxAnimation['shapeFilter'] | undefined) ?? presetShape,
                    shapeDirection: (shapeFilter?.[2] as PptxAnimation['shapeDirection'] | undefined) ?? presetShapeDirection,
                    slideFrom: (slide?.[1]?.toLowerCase() as PptxAnimation['slideFrom'] | undefined) || presetFlyDirection || presetCrawlDirection || presetPeekDirection || (presetRiseUp ? 'bottom' : undefined),
                    wheelSpokes: wheel ? (Number(wheel[1]) as PptxAnimation['wheelSpokes']) : presetWheelSpokes
                })
                handledEffects.add(effect)
            } else {
                warnings.add(`The ${transition || 'unknown'} ${filter || 'unknown'} object animation is not rendered.`)
                preserveEffectTiming()
            }
        }
        for (const animation of ownedDescendants('animScale')) {
            const animationTarget = target(animation)
            const timeNode = firstDescendant(animation, 'cTn')
            const element = elementsById.get(animationTarget.targetId)
            const timing = animationTiming(timeNode, 'Scale')
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                warnings.add('A scale animation with an unsupported target or text range was skipped.')
                preserveTimingOnly(animationTarget.targetId, timeNode, timing)
                handledScales.add(animation)
                continue
            }
            if (!timing) {
                handledScales.add(animation)
                continue
            }
            // DrawingML scale coordinates are in 100,000ths; the SDK example uses by=150000 for a 150% scale.
            // https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.presentation.animatescale?view=openxml-3.0.1
            const pair = (node: Element): [number, number] | undefined => {
                const x = percentageValue(node.getAttribute('x'))
                const y = percentageValue(node.getAttribute('y'))
                return x !== undefined && y !== undefined ? [x, y] : undefined
            }
            const fromNode = child(animation, 'from')
            const toNode = child(animation, 'to')
            const byNode = child(animation, 'by')
            const scaleBehavior = child(animation, 'cBhvr')
            const scaleAttributeList = child(scaleBehavior, 'attrNameLst')
            const scaleAttributeNames = descendants(scaleAttributeList, 'attrName')
                .map(attribute => normalizeAnimationAttributeName(attribute.textContent))
            // PowerPoint requires ScaleX and ScaleY for animScale; if a list is present,
            // don't reinterpret other explicitly targeted properties as an x/y scale pair.
            // Keep the existing fallback for older files that omit attrNameLst entirely.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/358c6de3-2cd7-459b-8eb9-b90f3721fd12
            const hasValidScaleAttributes = !scaleAttributeList || (
                scaleAttributeNames.length >= 2 &&
                scaleAttributeNames[0] === 'ScaleX' &&
                scaleAttributeNames[1] === 'ScaleY'
            )
            if (!hasValidScaleAttributes) {
                warnings.add('A scale animation without the required ScaleX and ScaleY attributes was skipped.')
                preserveTimingOnly(animationTarget.targetId, timeNode, timing)
                handledScales.add(animation)
                continue
            }
            const scaleAdditive = scaleBehavior?.getAttribute('additive')
            if (scaleAdditive !== null && scaleAdditive !== undefined && !['base', 'sum', 'repl', 'mult', 'none'].includes(scaleAdditive)) {
                warnings.add('A scale animation with an unsupported additive mode was skipped.')
                preserveTimingOnly(animationTarget.targetId, timeNode, timing)
                handledScales.add(animation)
                continue
            }
            // OOXML's raw to-only scale default is (0, 0), while by-only starts at (1, 1).
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/358c6de3-2cd7-459b-8eb9-b90f3721fd12
            const explicitFrom = fromNode ? pair(fromNode) : undefined
            const explicitTo = toNode ? pair(toNode) : undefined
            // Office ignores by when to is present for an animScale from_to/to_only track.
            const by = byNode && !toNode ? pair(byNode) : undefined
            const hasInvalidScaleCoordinate = Boolean(
                (fromNode && !explicitFrom) ||
                (toNode && !explicitTo) ||
                (byNode && !toNode && !by)
            )
            const scaleFrom = explicitFrom ?? (explicitTo ? ([0, 0] as [number, number]) : ([1, 1] as [number, number]))
            const scaleTo = explicitTo ?? (by ? ([scaleFrom[0] * by[0], scaleFrom[1] * by[1]] as [number, number]) : undefined)
            // MS-OI29500 permits scale percentages up to 2,147,483,625 thousandths.
            const maximumScaleFactor = 2_147_483_625 / 100_000
            const encodedScaleValues = [...(explicitFrom || []), ...(explicitTo || []), ...(by || [])]
            if (hasInvalidScaleCoordinate || !scaleTo || encodedScaleValues.some(value => !Number.isFinite(value) || value < 0 || value > maximumScaleFactor) || scaleTo.some(value => !Number.isFinite(value) || value < 0)) {
                warnings.add('A scale animation with missing or out-of-range values was skipped.')
                preserveTimingOnly(animationTarget.targetId, timeNode, timing)
                handledScales.add(animation)
                continue
            }
            // In this composition layer, base/repl/none replace, while sum and mult compose with other scale animations.
            const key = `${animationTarget.targetId}:scale:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'scale',
                direction: 'in',
                scaleFrom,
                scaleTo,
                scaleAdditive: scaleAdditive === 'sum',
                scaleMultiplicative: scaleAdditive === 'mult',
                scaleAccumulate: child(animation, 'cBhvr')?.getAttribute('accumulate') === 'always',
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + resolvedStartDelay(timeNode),
                ...timing
            })
            handledScales.add(animation)
        }
        for (const animation of ownedDescendants('animRot')) {
            const behavior = child(animation, 'cBhvr')
            const animationTarget = target(behavior || animation)
            const element = elementsById.get(animationTarget.targetId)
            const timeNode = child(behavior, 'cTn') || firstDescendant(animation, 'cTn')
            const timing = animationTiming(timeNode, 'Rotation')
            const names = descendants(child(behavior, 'attrNameLst'), 'attrName')
            const property = normalizeAnimationAttributeName(names[0]?.textContent)
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element) || !['r', 'ppt_r', 'style.rotation'].includes(property || '')) {
                warnings.add('A rotation animation with an unsupported target or property was skipped.')
                preserveTimingOnly(animationTarget.targetId, timeNode, timing)
                handledRotations.add(animation)
                continue
            }
            // DrawingML ST_Angle values are measured in 60,000ths of a degree.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animaterotation?view=openxml-3.0.1
            // ST_Angle maps to a signed Int32Value, including -2^31 as its minimum.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rotation?view=openxml-3.0.1
            const rawFrom = animation.getAttribute('from')
            const rawTo = animation.getAttribute('to')
            const rawBy = animation.getAttribute('by')
            const parseAngle = (value: string | null) => {
                const lexicalValue = trimXmlWhitespace(value)
                if (lexicalValue === undefined || !/^[-+]?\d+$/.test(lexicalValue)) return undefined
                const angle = Number(lexicalValue)
                return Number.isInteger(angle) && angle >= -2_147_483_648 && angle <= 2_147_483_647 ? angle / 60_000 : undefined
            }
            const from = rawFrom === null ? undefined : parseAngle(rawFrom)
            const to = rawTo === null ? undefined : parseAngle(rawTo)
            const by = rawBy === null ? undefined : parseAngle(rawBy)
            const valid = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                && (rawFrom === null || from !== undefined)
                && (rawTo !== null ? to !== undefined : by !== undefined)
            if (!valid) {
                warnings.add('A rotation animation with an invalid from/to/by combination was skipped.')
                preserveTimingOnly(animationTarget.targetId, timeNode, timing)
                handledRotations.add(animation)
                continue
            }
            // MS-OI29500 §19.5.5(b): to_only and by_only rotation tracks start at 0.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/39903f1a-fcb4-4aee-81e2-66f4fc7493fd
            const rotationFrom = from ?? 0
            const rotationTo = to ?? rotationFrom + (by ?? 0)
            const rotationAdditive = behavior?.getAttribute('additive')
            if (!timing) {
                handledRotations.add(animation)
                continue
            }
            const key = `${animationTarget.targetId}:rotation:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'rotation',
                direction: 'in',
                rotationFrom,
                rotationTo,
                rotationRelative: from === undefined && to === undefined,
                rotationAdditive: rotationAdditive === 'sum',
                rotationAccumulate: behavior?.getAttribute('accumulate') === 'always',
                durationMs: animationDurationMs(timeNode, timing),
                delayMs: delayMs + resolvedStartDelay(timeNode),
                ...timing
            })
            handledRotations.add(animation)
        }
        // Grow With Color writes style tracks that start from the current text style,
        // even though its exported to-only values omit <p:from>.
        // Require the complete observed signature so other iterated color effects
        // keep the documented defaults for omitted <p:from> values.
        const ownedColorTracks = ownedDescendants('animClr')
        const ownedGenericTracks = ownedDescendants('anim')
        const ownedSetTracks = ownedDescendants('set')
        const trackProperty = (track: Element) => normalizeAnimationAttributeName(descendants(child(child(track, 'cBhvr'), 'attrNameLst'), 'attrName')[0]?.textContent)
        const growWithColorStyleColorTrack = ownedColorTracks.find(track => trackProperty(track) === 'style.color'
            && !child(track, 'from') && Boolean(child(track, 'to'))
            && child(track, 'cBhvr')?.getAttribute('override') === 'childStyle')
        const growWithColorFillColorTrack = ownedColorTracks.find(track => trackProperty(track) === 'fillcolor'
            && !child(track, 'from') && Boolean(child(track, 'to')))
        const growWithColorFillTypeTrack = ownedSetTracks.find(track => trackProperty(track) === 'fill.type'
            && firstDescendant(child(track, 'to'), 'strVal')?.getAttribute('val') === 'solid')
        const growWithColorFontSizeTrack = ownedGenericTracks.find(track => trackProperty(track) === 'style.fontSize'
            && track.getAttribute('to') === '1.5' && !track.hasAttribute('from') && !track.hasAttribute('by')
            && child(track, 'cBhvr')?.getAttribute('override') === 'childStyle')
        const growWithColorTrackTargetIds = [growWithColorStyleColorTrack, growWithColorFillColorTrack, growWithColorFillTypeTrack, growWithColorFontSizeTrack]
            .map(track => track ? target(child(track, 'cBhvr') || track).targetId : '')
        const hasGrowWithColorSignature = owner?.getAttribute('nodeType') === 'clickEffect'
            && trimXmlWhitespace(owner.getAttribute('presetID')) === '28'
            && trimXmlWhitespace(owner.getAttribute('presetClass')) === 'emph'
            && trimXmlWhitespace(owner.getAttribute('presetSubtype')) === '0'
            && iterate?.getAttribute('type') === 'lt'
            && child(iterate, 'tmPct')?.getAttribute('val') === '10000'
            && ownedColorTracks.length === 2
            && ownedGenericTracks.length === 1
            && ownedSetTracks.length === 1
            && Boolean(growWithColorStyleColorTrack && growWithColorFillColorTrack && growWithColorFillTypeTrack && growWithColorFontSizeTrack)
            && Boolean(growWithColorTrackTargetIds[0])
            && growWithColorTrackTargetIds.every(targetId => targetId === growWithColorTrackTargetIds[0])
        const isWpsGrowWithColor = isWpsApplication && hasGrowWithColorSignature
        const isGrowWithColorStyleColorTrack = hasGrowWithColorSignature && growWithColorStyleColorTrack !== undefined
        const isGrowWithColorFontSizeTrack = hasGrowWithColorSignature && growWithColorFontSizeTrack !== undefined
        // WPS playback completes the object-fill tween before starting the
        // per-letter childStyle color and font-size tracks.
        const wpsGrowWithColorChildStyleDelayMs = isWpsGrowWithColor
            ? timelineTime(child(child(growWithColorFillColorTrack, 'cBhvr'), 'cTn')?.getAttribute('dur'), 0)
            : 0
        const relativeColorProperties = ['fillcolor', 'stroke.color', 'style.color']
        const wpsRelativeColorTracks = ownedColorTracks.filter(track =>
            !child(track, 'from') && !child(track, 'to') && Boolean(child(track, 'by'))
        )
        const relativeColorByValue = (track: Element) => {
            const value = children(child(track, 'by'))[0]
            return value?.localName === 'hsl'
                ? ['h', 's', 'l'].map(attribute => value.getAttribute(attribute) || '').join(':')
                : ''
        }
        // WPS exports Lighten (30) and Desaturate (25) as a matched set of
        // relative HSL tracks. Resolve those deltas against each property's
        // saved color, while leaving ordinary omitted-from animClr at black.
        const wpsCurrentColorEmphasis = owner?.getAttribute('nodeType') === 'clickEffect'
            && trimXmlWhitespace(owner.getAttribute('presetClass')) === 'emph'
            && trimXmlWhitespace(owner.getAttribute('presetSubtype')) === '0'
            && ['25', '30'].includes(trimXmlWhitespace(owner.getAttribute('presetID')) || '')
            && !iterate
            && ownedColorTracks.length === 3
            && ownedSetTracks.length === 1
            && wpsRelativeColorTracks.length === 3
            && wpsRelativeColorTracks.every(track => track.getAttribute('clrSpc') === 'hsl')
            && relativeColorProperties.every(property => wpsRelativeColorTracks.some(track => trackProperty(track) === property))
            && wpsRelativeColorTracks.some(track => trackProperty(track) === 'style.color' && child(track, 'cBhvr')?.getAttribute('override') === 'childStyle')
            && wpsRelativeColorTracks.every(track => target(child(track, 'cBhvr') || track).targetId === target(child(wpsRelativeColorTracks[0]!, 'cBhvr') || wpsRelativeColorTracks[0]!).targetId)
            && Boolean(target(child(wpsRelativeColorTracks[0]!, 'cBhvr') || wpsRelativeColorTracks[0]!).targetId)
            && wpsRelativeColorTracks.every(track => relativeColorByValue(track) !== '')
            && new Set(wpsRelativeColorTracks.map(relativeColorByValue)).size === 1
            && trackProperty(ownedSetTracks[0]!) === 'fill.type'
            && firstDescendant(child(ownedSetTracks[0]!, 'to'), 'strVal')?.getAttribute('val') === 'solid'
            && target(child(ownedSetTracks[0]!, 'cBhvr') || ownedSetTracks[0]!).targetId === target(child(wpsRelativeColorTracks[0]!, 'cBhvr') || wpsRelativeColorTracks[0]!).targetId
        for (const animation of ownedColorTracks) {
            const behavior = child(animation, 'cBhvr')
            const animationTarget = target(behavior || animation)
            const property = normalizeAnimationAttributeName(descendants(child(behavior, 'attrNameLst'), 'attrName')[0]?.textContent)
            if (property === 'imageData.chromakey') {
                // PowerPoint limits p:animClr targets to style.color, fillcolor, and stroke.color.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/9636f7f3-58e3-408d-9880-d0799fb26b0f
                warnings.add('PowerPoint does not support imageData.chromakey as a p:animClr target; the animation was skipped.')
                handledColors.add(animation)
                continue
            }
            // The real WPS Grow With Color preset only animates the shape fill.
            // Its childStyle text-color track is present in OOXML but WPS leaves
            // the saved text color unchanged during playback.
            const element = elementsById.get(animationTarget.targetId)
            const rawColorSpace = animation.getAttribute('clrSpc')
            const rawColorDirection = animation.getAttribute('dir')
            // Apply defaults only when omitted; present values must match the schema enums.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/9636f7f3-58e3-408d-9880-d0799fb26b0f
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animatecolorspacevalues?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animatecolordirectionvalues?view=openxml-3.0.1
            const colorSpace = rawColorSpace === null ? 'rgb' : trimXmlWhitespace(rawColorSpace) ?? ''
            const colorDirection = rawColorDirection === null ? 'cw' : trimXmlWhitespace(rawColorDirection) ?? ''
            const invalidTarget =
                !element ||
                animationTarget.hasUnsupportedTextRange ||
                ((animationTarget.paragraphRange || animationTarget.characterRange) && !['style.color', 'stroke.color'].includes(property || '')) ||
                !hasValidTextRange(animationTarget, element) ||
                (property === 'fillcolor' && (element?.kind !== 'shape' || /gradient\(/i.test(element.fill))) ||
                (property === 'stroke.color' && (!['shape', 'line'].includes(element?.kind || '') || /gradient\(/i.test(element.stroke))) ||
                (property === 'style.color' && (!element?.paragraphs.length || (element.kind === 'table' && animationTarget.paragraphRange)))
            if (!animationTarget.targetId || invalidTarget || !['fillcolor', 'style.color', 'stroke.color'].includes(property || '')) {
                warnings.add('A color animation with an unsupported target or property was skipped.')
                handledColors.add(animation)
                continue
            }
            const usesCurrentColor = (wpsCurrentColorEmphasis
                && !child(animation, 'from')
                && !child(animation, 'to')
                && Boolean(child(animation, 'by'))
                && ['fillcolor', 'style.color', 'stroke.color'].includes(property || ''))
                || (isGrowWithColorStyleColorTrack && animation === growWithColorStyleColorTrack && property === 'style.color')
            const hasTextTarget = Boolean(animationTarget.paragraphRange || animationTarget.characterRange)
            const currentColorBaseline = property === 'fillcolor' ? element.fill : property === 'stroke.color' && !hasTextTarget ? element.stroke : undefined
            if (usesCurrentColor && property !== 'style.color' && !(property === 'stroke.color' && hasTextTarget)) {
                const channels = animationColorStringChannels(currentColorBaseline, 'hsl')
                if (!channels || channels[3] <= 0) {
                    warnings.add(`A WPS relative color emphasis track targeting an unset ${property === 'fillcolor' ? 'shape fill' : 'stroke'} was skipped.`)
                    handledColors.add(animation)
                    continue
                }
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
            // Grow With Color uses to-only fill and childStyle color tracks;
            // its native playback starts each track at the saved property value.
            // Keep the documented black default for other omitted-from color tracks.
            const iteratedFillColor = !fromNode
                && property === 'fillcolor'
                && hasGrowWithColorSignature
                && element?.kind === 'shape'
                ? animationColorStringChannels(element.fill, colorSpace as 'rgb' | 'hsl')
                : undefined
            const colorFrom = fromNode ? colorValue(fromNode) : iteratedFillColor ?? black
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
            const key = `${animationTarget.targetId}:color:${property}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'color',
                direction: 'in',
                colorProperty: property as PptxAnimation['colorProperty'],
                colorFrom,
                colorTo,
                ...(usesCurrentColor ? {...(colorBy ? {colorBy} : {}), colorFromCurrent: true} : {}),
                colorSpace: colorSpace as PptxAnimation['colorSpace'],
                colorDirection: colorDirection as PptxAnimation['colorDirection'],
                durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + (isWpsGrowWithColor && animation === growWithColorStyleColorTrack ? wpsGrowWithColorChildStyleDelayMs : 0) + resolvedStartDelay(timeNode),
                ...timing
            })
            handledColors.add(animation)
        }
        for (const animation of ownedDescendants('anim')) {
            // Microsoft lists style.fontWeight among p:tav formula-capable target attributes.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa
            // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
            // The Open XML SDK specifies that p:tav formulas support both linear and discrete calculation modes.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timeanimatevalue.fomula?view=openxml-3.0.1
            // Preserve the declared linear or discrete mode when sampling formula-based numeric p:tav tracks.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timeanimatevalue?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animatebehaviorcalculatemodevalues?view=openxml-3.0.1
            const behavior = child(animation, 'cBhvr')
            const animationTarget = target(behavior || animation)
            const rawProperty = descendants(child(behavior, 'attrNameLst'), 'attrName')[0]?.textContent
            const property = normalizeAnimationAttributeName(rawProperty)
            const isGrowWithColorFontSize = isGrowWithColorFontSizeTrack && animation === growWithColorFontSizeTrack && property === 'style.fontSize'
            const element = elementsById.get(animationTarget.targetId)
            const timeNode = child(behavior, 'cTn') || firstDescendant(animation, 'cTn')
            const tavList = child(animation, 'tavLst')
            const keyframeNodes = tavList ? children(tavList) : []
            const hasNumericFormula = keyframeNodes.some(node => Boolean(node.getAttribute('fmla')?.trim()))
            let calculationMode = animation.getAttribute('calcmode') || 'lin'
            // p:tav formulas support linear and discrete interpolation. When a formula track
            // declares another mode, ECMA-376 says to assume linear interpolation.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timeanimatevalue.fomula?view=openxml-3.0.1
            if (hasNumericFormula && calculationMode !== 'lin' && calculationMode !== 'discrete') calculationMode = 'lin'
            const valueType = animation.getAttribute('valueType') || 'num'
            const textStyleProperty = property === 'style.fontWeight' && animation.getAttribute('valueType') === 'str' ? 'fontWeight' : property === 'style.fontStyle' ? 'fontStyle' : property === 'style.fontFamily' ? 'fontFamily' : property === 'style.textDecorationUnderline' ? 'underline' : property === 'style.textDecorationLineThrough' ? 'lineThrough' : property === 'style.textTransform' ? 'textTransform' : property === 'style.textEffectOutline' ? 'outline' : property === 'style.textShadow' ? 'textShadow' : property === 'style.textEffectEmboss' ? 'emboss' : undefined
            const isColorProperty = ['fillcolor', 'style.color', 'stroke.color', 'shadow.color'].includes(property || '')
            const isVisibilityProperty = property === 'style.visibility'
            const isGrayscaleProperty = property === 'imageData.grayscale'
            const isOnOffProperty = property === 'fill.on' || property === 'stroke.on' || property === 'shadow.on'
            const isStrokeDashProperty = property === 'stroke.dashstyle'
            const isLineArrowTypeProperty = property === 'stroke.startArrow' || property === 'stroke.endArrow'
            const isLineArrowSizeProperty = property === 'stroke.startArrowWidth' || property === 'stroke.endArrowWidth' || property === 'stroke.startArrowLength' || property === 'stroke.endArrowLength'
            const isLineArrowProperty = isLineArrowTypeProperty || isLineArrowSizeProperty
            const isShadowOffsetProperty = property === 'shadow.offset.x' || property === 'shadow.offset.y'
            const isCoordinateProperty = property === 'ppt_x' || property === 'ppt_y'
            const isDimensionProperty = property === 'ppt_w' || property === 'ppt_h'
            const isScaleProperty = property === 'ScaleX' || property === 'ScaleY' || isDimensionProperty
            const imageCropProperty = property === 'imageData.cropLeft' ? 'left'
                : property === 'imageData.cropTop' ? 'top'
                : property === 'imageData.cropRight' ? 'right'
                : property === 'imageData.cropBottom' ? 'bottom'
                : undefined
            const imageToneProperty = property === 'imageData.gain' ? 'gain'
                : property === 'imageData.blacklevel' || property === 'imageData.blackleve' ? 'blacklevel'
                : property === 'imageData.gamma' ? 'gamma'
                : undefined
            const hasTextRange = Boolean(animationTarget.paragraphRange || animationTarget.characterRange)
            const isRotationProperty = property === 'style.rotation'
            const directPresetChildren = owner ? children(child(owner, 'childTnLst')) : []
            const visibilitySet = directPresetChildren.find(node => node.localName === 'set')
            const visibilitySetBehavior = visibilitySet ? child(visibilitySet, 'cBhvr') : undefined
            const visibilitySetTimeNode = visibilitySetBehavior ? child(visibilitySetBehavior, 'cTn') : undefined
            const visibilitySetTarget = visibilitySetBehavior ? target(visibilitySetBehavior) : undefined
            const visibilitySetProperty = firstDescendant(child(visibilitySetBehavior, 'attrNameLst'), 'attrName')?.textContent?.trim()
            const visibilitySetValue = child(child(visibilitySet, 'to'), 'strVal')?.getAttribute('val')
            const isRandomEffectPlaceholder = Boolean(
                // Match the observed preset-24 placeholder exactly; other incomplete p:anim tracks stay unsupported.
                owner?.getAttribute('nodeType') === 'clickEffect' &&
                trimXmlWhitespace(owner.getAttribute('presetID')) === '24' &&
                trimXmlWhitespace(owner.getAttribute('presetClass')) === 'entr' &&
                trimXmlWhitespace(owner.getAttribute('presetSubtype')) === '0' &&
                directPresetChildren.length === 2 &&
                directPresetChildren.filter(node => node.localName === 'anim').length === 1 &&
                directPresetChildren.filter(node => node.localName === 'set').length === 1 &&
                ownedDescendants('anim').length === 1 &&
                ownedDescendants('animEffect').length === 0 &&
                visibilitySetProperty === 'style.visibility' &&
                visibilitySetTarget?.targetId === animationTarget.targetId &&
                visibilitySetValue === 'visible' &&
                !rawProperty &&
                animation.getAttribute('to') === '' &&
                !animation.hasAttribute('from') &&
                !animation.hasAttribute('by') &&
                !tavList &&
                calculationMode === 'lin' &&
                valueType === 'num' &&
                Boolean(timeNode) &&
                Boolean(element) &&
                !hasTextRange &&
                !animationTarget.hasUnsupportedTextRange &&
                hasValidTextRange(animationTarget, element)
            )
            if (isRandomEffectPlaceholder) {
                const timing = animationTiming(timeNode, 'Random Effects')
                handledTextAnimations.add(animation)
                if (!timing) continue
                const variant = randomEffectVariant()
                // These placeholder tracks can be 1ms even when the paired visibility restore marks a 2s effect.
                const visibilitySetDurationMs = visibilitySetTimeNode
                    ? resolvedStartDelay(visibilitySetTimeNode) + animationDurationMs(visibilitySetTimeNode, timing, 1)
                    : 0
                actions.set(`${animationTarget.targetId}:random-preset:${timeNode?.getAttribute('id') || actions.size}`, {
                    targetId: animationTarget.targetId,
                    ...variant,
                    randomEffectFallback: true,
                    direction: 'in',
                    durationMs: Math.max(animationDurationMs(timeNode, timing), visibilitySetDurationMs),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                warnings.add(`This Random Effects preset omitted its selected effect; ${variant.effect} was randomly selected for this playback session.`)
                continue
            }
            if (
                !animationTarget.targetId ||
                !['style.fontSize', 'style.fontWeight', 'style.fontStyle', 'style.fontFamily', 'style.textDecorationUnderline', 'style.textDecorationLineThrough', 'style.textTransform', 'style.textEffectOutline', 'style.textShadow', 'style.textEffectEmboss', 'style.visibility', 'style.rotation', 'imageData.grayscale', 'fill.on', 'fillcolor', 'style.color', 'stroke.on', 'stroke.color', 'shadow.on', 'shadow.color', 'shadow.offset.x', 'shadow.offset.y', 'stroke.dashstyle', 'stroke.startArrow', 'stroke.endArrow', 'stroke.startArrowWidth', 'stroke.endArrowWidth', 'stroke.startArrowLength', 'stroke.endArrowLength', 'imageData.cropLeft', 'imageData.cropTop', 'imageData.cropRight', 'imageData.cropBottom', 'imageData.gain', 'imageData.blacklevel', 'imageData.blackleve', 'imageData.gamma', 'style.opacity', 'fill.opacity', 'stroke.opacity', 'stroke.weight', 'shadow.opacity', 'ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ScaleX', 'ScaleY'].includes(property || '') ||
                !element ||
                (isCoordinateProperty && hasTextRange) ||
                (isScaleProperty && hasTextRange && !isDimensionProperty) ||
                (isOnOffProperty && hasTextRange) ||
                (isStrokeDashProperty && (!['shape', 'line'].includes(element.kind) || hasTextRange)) ||
                (isLineArrowProperty && (element.kind !== 'line' || hasTextRange)) ||
                (property === 'stroke.weight' && (!['shape', 'line'].includes(element.kind) || hasTextRange)) ||
                (['style.fontSize', 'style.fontWeight', 'style.fontStyle', 'style.fontFamily', 'style.textDecorationUnderline', 'style.textDecorationLineThrough', 'style.textTransform', 'style.textEffectOutline', 'style.textShadow', 'style.textEffectEmboss'].includes(property || '') && !element.paragraphs.length) ||
                (isGrayscaleProperty && (element.kind !== 'picture' || hasTextRange)) ||
                (imageCropProperty && (element.kind !== 'picture' || hasTextRange)) ||
                (imageToneProperty && (element.kind !== 'picture' || hasTextRange)) ||
                (hasTextRange && isColorProperty && !['style.color', 'shadow.color'].includes(property || '')) ||
                (property === 'fillcolor' && (element.kind !== 'shape' || /gradient\(/i.test(element.fill))) ||
                (property === 'stroke.color' && (!['shape', 'line'].includes(element.kind) || /gradient\(/i.test(element.stroke))) ||
                (property === 'shadow.color' && !hasShadow(element)) ||
                (isShadowOffsetProperty && !hasShadow(element)) ||
                (property === 'style.color' && (!element.paragraphs.length || (element.kind === 'table' && animationTarget.paragraphRange))) ||
                (property === 'fill.opacity' && !hasTextRange && element.kind !== 'shape') ||
                (property === 'stroke.opacity' && !hasTextRange && !supportsStrokeOpacity(element)) ||
                (property === 'shadow.opacity' && !hasTextRange && !hasShadow(element)) ||
                animationTarget.hasUnsupportedTextRange ||
                !hasValidTextRange(animationTarget, element) ||
                valueType !== (isColorProperty ? 'clr' : textStyleProperty || isVisibilityProperty || isGrayscaleProperty || isOnOffProperty || isStrokeDashProperty || isLineArrowProperty ? 'str' : 'num') ||
                !['lin', 'discrete'].includes(calculationMode) ||
                (calculationMode === 'discrete' && !tavList && !textStyleProperty && !isVisibilityProperty && !isGrayscaleProperty && !isOnOffProperty && !isStrokeDashProperty && !isLineArrowProperty && !isCoordinateProperty && !isScaleProperty && !isRotationProperty && !isShadowOffsetProperty && !isColorProperty) ||
                (tavList && children(tavList).some(node => node.localName !== 'tav'))
            ) {
                warnings.add('A generic object animation with an unsupported target or interpolation was skipped.')
                handledTextAnimations.add(animation)
                continue
            }
            const rawFrom = animation.getAttribute('from')
            const rawTo = animation.getAttribute('to')
            const rawBy = animation.getAttribute('by')
            const formulaGeometry = animationGeometryVariables(element, slideWidth, slideHeight)
            // Office/WPS dimension tracks can store #ppt_w/#ppt_h expressions in strVal.
            // MS-OE376 attrName and tav define these properties and shape-size variables.
            const parseNumericValue = (raw: string | null): number => {
                if (!raw?.trim() || isDimensionProperty && raw.length > 16_384) return Number.NaN
                const numeric = parseOfficeAnimationNumber(raw)
                if (numeric !== undefined || !isDimensionProperty) return numeric ?? Number.NaN
                return evaluateMotionFormula(raw, formulaGeometry, true) ?? Number.NaN
            }
            const parseFixedPercentage = (raw: string | null): number | undefined => {
                if (raw === null) return undefined
                const value = raw.trim()
                // ST_PositiveFixedPercentage accepts a trailing percent or 1000ths of a percent; Office writes the latter.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c1f1feac-e34c-48d5-b2e2-f67bb67113e7
                if (/^\d+(?:\.\d+)?%$/.test(value)) return Number(value.slice(0, -1)) / 100
                if (/^\d+$/.test(value)) return Number(value) / 100_000
                return undefined
            }
            const minimumKeyframeValue = isRotationProperty ? -2_147_483_647 : isDimensionProperty ? -1000 : isShadowOffsetProperty ? -MAX_ANIMATED_SHADOW_OFFSET_EMU : 0
            const maximumKeyframeValue = isRotationProperty ? 2_147_483_647 : property === 'stroke.weight' ? MAX_ANIMATED_STROKE_WIDTH_EMU : isShadowOffsetProperty ? MAX_ANIMATED_SHADOW_OFFSET_EMU : 1000
            // ponytail: cap at 256 keypoints to bound parser and per-frame work; raise only when real decks need denser curves.
            const parsedKeyframes =
                tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes
                          .map(node => {
                              const valueNode = children(child(node, 'val'))[0]
                              const offset = parseFixedPercentage(node.getAttribute('tm'))
                              const rawValue = valueNode && ['fltVal', 'strVal'].includes(valueNode.localName) ? valueNode.getAttribute('val') : null
                              const value = parseNumericValue(rawValue)
                              const formula = node.getAttribute('fmla')?.trim() || undefined
                              return offset !== undefined && offset >= 0 && offset <= 1 && Number.isFinite(value) && (hasNumericFormula || value >= minimumKeyframeValue && value <= maximumKeyframeValue)
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
            if (isColorProperty) {
                // clr tracks use RGB/HSL strings or p:clrVal keypoints; a to-only or by-only track starts at black.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/cd27d302-0342-4d80-8118-5495c8fafea8
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.progress?view=openxml-3.0.1
                const keyframeColorSpaces = keyframeNodes.map(node => {
                    const valueNode = children(child(node, 'val'))[0]
                    if (valueNode?.localName === 'clrVal') return 'rgb' as const
                    if (valueNode?.localName === 'strVal') return parseAnimationColorString(valueNode.getAttribute('val'))?.colorSpace
                    return undefined
                })
                const directColorSpaces = [rawFrom, rawTo, rawTo === null ? rawBy : null]
                    .filter((value): value is string => value !== null)
                    .map(value => parseAnimationColorString(value)?.colorSpace)
                const colorSpaces = [...directColorSpaces, ...keyframeColorSpaces].filter((space): space is AnimationColorSpace => Boolean(space))
                const colorSpace: AnimationColorSpace = colorSpaces.length > 0 && colorSpaces.every(space => space === 'hsl') ? 'hsl' : 'rgb'
                const parsedColorKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map(node => {
                          const offset = parseFixedPercentage(node.getAttribute('tm'))
                          const valueNode = children(child(node, 'val'))[0]
                          const value = valueNode?.localName === 'strVal'
                              ? animationColorStringChannels(valueNode.getAttribute('val'), colorSpace)
                              : valueNode?.localName === 'clrVal'
                                ? animationColorChannels(children(valueNode)[0], theme, colorSpace)
                                : undefined
                          const formula = node.getAttribute('fmla')?.trim()
                          return offset !== undefined && offset >= 0 && offset <= 1 && value && !formula
                              ? {offset, value}
                              : undefined
                      })
                    : undefined
                const validColorKeyframes = Boolean(parsedColorKeyframes
                    && parsedColorKeyframes.length > 0
                    && parsedColorKeyframes.length === keyframeNodes.length
                    && parsedColorKeyframes.every((frame, index) => {
                        if (!frame) return false
                        if (index === 0) return true
                        const previous = parsedColorKeyframes[index - 1]
                        return Boolean(previous && frame.offset > previous.offset)
                    }))
                const colorKeyframes = validColorKeyframes
                    ? parsedColorKeyframes!.map(frame => frame!)
                    : undefined
                const hasFrom = rawFrom !== null
                const hasTo = rawTo !== null
                const hasBy = rawBy !== null
                const validDirectCombination = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const validValues = tavList
                    ? validColorKeyframes && !hasFrom && !hasTo && !hasBy
                    : validDirectCombination
                const black: [number, number, number, number] = [0, 0, 0, 1]
                const colorFrom = colorKeyframes?.[0]?.value ?? (hasFrom ? animationColorStringChannels(rawFrom, colorSpace) : black)
                const explicitTo = colorKeyframes?.at(-1)?.value ?? (hasTo ? animationColorStringChannels(rawTo, colorSpace) : undefined)
                const colorBy = hasBy && !hasTo ? animationColorStringChannels(rawBy, colorSpace) : undefined
                const colorTo = explicitTo || (colorFrom && colorBy
                    ? colorFrom.map((channel, index) => {
                          const maximum = index === 3 ? 1 : colorSpace === 'rgb' ? 255 : 1
                          const sum = channel + colorBy[index]!
                          return index === 0 && colorSpace === 'hsl' ? sum : Math.max(0, Math.min(maximum, sum))
                      }) as [number, number, number, number]
                    : undefined)
                const additive = behavior?.getAttribute('additive')
                if (!validValues || !colorFrom || !colorTo || (additive !== null && additive !== undefined && !['base', 'sum'].includes(additive))) {
                    warnings.add('A generic color animation with missing, invalid, or unsupported values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Color')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const directColorKeyframes = !tavList && calculationMode === 'discrete'
                    ? [{offset: 0, value: colorFrom}, {offset: 1, value: colorTo}]
                    : undefined
                actions.set(`${animationTarget.targetId}:color:generic:${property}:${timeNode?.getAttribute('id') || actions.size}`, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'color',
                    direction: 'in',
                    colorProperty: property as PptxAnimation['colorProperty'],
                    colorFrom,
                    colorTo,
                    colorKeyframes: colorKeyframes || directColorKeyframes,
                    colorKeyframeMode: tavList || directColorKeyframes ? calculationMode as 'lin' | 'discrete' : undefined,
                    colorSpace,
                    colorDirection: 'cw',
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isRotationProperty) {
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                const validKeyframeTrack = Boolean(parsedKeyframes?.length && parsedKeyframes.length === keyframeNodes.length && parsedKeyframes.every((frame, index) => index === 0 || frame.offset > parsedKeyframes[index - 1]!.offset))
                const validRotationKeyframes = validKeyframeTrack
                const validCombination = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const from = parsedKeyframes?.[0]?.value ?? (rawFrom !== null ? parseNumericValue(rawFrom) : rawTo !== null || rawBy !== null ? 0 : Number.NaN)
                const to = parsedKeyframes?.at(-1)?.value ?? (rawTo !== null ? parseNumericValue(rawTo) : rawBy !== null ? from + parseNumericValue(rawBy) : undefined)
                const additive = behavior?.getAttribute('additive')
                const validValues = tavList
                    ? validRotationKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : validCombination
                if (!validValues || !Number.isFinite(from) || to === undefined || !Number.isFinite(to) || (!hasFormulaKeyframes && (from < -2_147_483_648 || from > 2_147_483_647 || to < -2_147_483_648 || to > 2_147_483_647)) || (additive !== null && additive !== undefined && !['base', 'sum'].includes(additive))) {
                    warnings.add(hasFormulaKeyframes
                        ? 'A generic style.rotation animation had an unsupported formula or keyframe track and was skipped.'
                        : 'A generic style.rotation animation with unsupported composition or out-of-range values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Rotation')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const formulaSamples = hasFormulaKeyframes && parsedKeyframes
                    ? sampleNumericAnimationFormula(parsedKeyframes, durationMs, 'style.rotation', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && !formulaSamples) {
                    warnings.add('A generic style.rotation formula used an unsupported variable or produced an out-of-range value; animation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                // WPS generic style.rotation tracks use plain degrees (for example, Grow and Turn exports 90 -> 0).
                // Preserve DrawingML angle units for aliases, formula tracks, and large values used by other exports.
                const rawRotationValues = [
                    rawFrom,
                    rawTo,
                    ...(rawTo === null ? [rawBy] : []),
                    ...(parsedKeyframes?.map(frame => String(frame.value)) ?? [])
                ].filter((value): value is string => value !== null).map(value => parseOfficeAnimationNumber(value) ?? Number.NaN)
                const usesCompactStyleRotationDegrees = rawProperty?.trim().toLowerCase() === 'style.rotation'
                    && !hasFormulaKeyframes
                    && rawRotationValues.length > 0
                    && rawRotationValues.every(value => Number.isFinite(value) && Math.abs(value) <= 360)
                const rotationUnitScale = usesCompactStyleRotationDegrees ? 1 : 1 / 60_000
                const rotationFrom = formulaSamples?.[0]?.value ?? from
                const rotationTo = formulaSamples?.at(-1)?.value ?? to
                const fromDegrees = rotationFrom * rotationUnitScale
                const toDegrees = rotationTo * rotationUnitScale
                const rotationKeyframes = formulaSamples?.map(({offset, value}) => ({offset, value: value / 60_000}))
                    ?? keyframes?.map(({offset, value}) => ({offset, value: value * rotationUnitScale}))
                    ?? (calculationMode === 'discrete' ? [{offset: 0, value: fromDegrees}, {offset: 1, value: toDegrees}] : undefined)
                actions.set(`${animationTarget.targetId}:rotation:generic:${timeNode?.getAttribute('id') || actions.size}`, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'rotation',
                    direction: 'in',
                    rotationFrom: fromDegrees,
                    rotationTo: toDegrees,
                    rotationKeyframes,
                    rotationKeyframeMode: formulaSamples ? calculationMode as 'lin' | 'discrete' : tavList || calculationMode === 'discrete' ? calculationMode as 'lin' | 'discrete' : undefined,
                    rotationRelative: rawFrom === null && rawTo === null && rawBy !== null,
                    rotationAdditive: additive === 'sum',
                    rotationAccumulate: behavior?.getAttribute('accumulate') === 'always',
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isScaleProperty) {
                const scaleProperty = property as 'ScaleX' | 'ScaleY' | 'ppt_w' | 'ppt_h'
                const scaleAxis = scaleProperty === 'ScaleX' || scaleProperty === 'ppt_w' ? 'x' : 'y'
                const baseline = isDimensionProperty ? scaleAxis === 'x' ? formulaGeometry.ppt_w : formulaGeometry.ppt_h : 1
                const minimumScaleValue = isDimensionProperty ? -baseline * 10 : 0
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                const validScaleKeyframes = Boolean(validKeyframeTrack && (hasFormulaKeyframes || parsedKeyframes?.every(frame => frame.value >= minimumScaleValue && frame.value <= baseline * 10)))
                const rawFrom = animation.getAttribute('from')
                const rawTo = animation.getAttribute('to')
                const rawBy = animation.getAttribute('by')
                // MS-OI29500 §19.5.1: numeric to-only and by-only animations start at 0.
                const from = parsedKeyframes?.[0]?.value ?? (rawFrom !== null ? parseNumericValue(rawFrom) : rawTo !== null || rawBy !== null ? 0 : baseline)
                const to = parsedKeyframes?.at(-1)?.value ?? (rawTo !== null ? parseNumericValue(rawTo) : rawBy !== null ? from + parseNumericValue(rawBy) : undefined)
                const validCombination = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const additive = behavior?.getAttribute('additive')
                const validValues = tavList
                    ? validScaleKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : validCombination
                if (!validValues || !Number.isFinite(baseline) || baseline <= 0 || !Number.isFinite(from) || !Number.isFinite(to) || to === undefined || (!hasFormulaKeyframes && (from < minimumScaleValue || from > baseline * 10 || to < minimumScaleValue || to > baseline * 10)) || additive !== null && additive !== undefined && !['base', 'sum', 'mult'].includes(additive)) {
                    warnings.add(`A generic ${property} animation with unsupported composition or out-of-range values was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Scale')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const formulaSamples = hasFormulaKeyframes && parsedKeyframes
                    ? sampleNumericAnimationFormula(parsedKeyframes, durationMs, scaleProperty, element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && !formulaSamples) {
                    warnings.add(`A generic ${property} formula used an unsupported variable or produced an out-of-range value; animation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const sourceKeyframes = formulaSamples || keyframes || (calculationMode === 'discrete' ? [{offset: 0, value: from}, {offset: 1, value: to}] : undefined)
                const scaleKeyframes = isDimensionProperty ? sourceKeyframes?.map(frame => ({offset: frame.offset, value: frame.value / baseline})) : sourceKeyframes
                const scaleFrom = (formulaSamples?.[0]?.value ?? from) / baseline
                const scaleTo = (formulaSamples?.at(-1)?.value ?? to) / baseline
                // Office defaults cBhvr@additive to base, and numeric by-only p:anim tracks start at zero.
                // Apply generic p:anim values against the identity scale to preserve the original scale as the base.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/b9a20a78-6401-4a8b-8ca9-b31ead409132
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                const baseAdditive = additive === null || additive === 'base'
                actions.set(`${animationTarget.targetId}:scale:${property}:${timeNode?.getAttribute('id') || actions.size}`, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'scale',
                    direction: 'in',
                    scaleLayer: isDimensionProperty ? 'dimensions' : 'transform',
                    scaleFrom: scaleAxis === 'x' ? [scaleFrom, 1] : [1, scaleFrom],
                    scaleTo: scaleAxis === 'x' ? [scaleTo, 1] : [1, scaleTo],
                    scaleAxis,
                    scaleAdditive: additive === 'sum',
                    scaleMultiplicative: additive === 'mult',
                    // ppt_w/ppt_h keyframes are geometry values normalized into
                    // the dimension layer; adding the underlying geometry again
                    // would make a 0 -> original-size zoom render as 1x -> 2x.
                    // Dimension by-only tracks carry a relative delta (for example WPS Flip's -2x width).
                    // Apply that delta to the original geometry; to-only dimension tracks remain absolute sizes.
                    scaleBaseAdditive: baseAdditive && (!isDimensionProperty || rawBy !== null && rawTo === null),
                    scaleKeyframes,
                    scaleKeyframeMode: tavList || calculationMode === 'discrete' ? calculationMode as 'lin' | 'discrete' : undefined,
                    scaleAccumulate: behavior?.getAttribute('accumulate') === 'always',
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isCoordinateProperty) {
                const baseline = formulaGeometry[property]
                const formulaVariables = formulaGeometry
                const readCoordinateValue = (raw: string | null, progress: number): number | undefined => {
                    const value = raw?.trim()
                    if (!value || value.length > 16_384) return undefined
                    const numeric = parseOfficeAnimationNumber(value)
                    if (numeric !== undefined) return numeric
                    return evaluateMotionFormula(value, {...formulaVariables, $: progress}, true)
                }
                const coordinateKeyframes = tavList && keyframeNodes.length >= 2 && keyframeNodes.length <= 256
                    ? keyframeNodes.map(node => {
                          const offset = parseFixedPercentage(node.getAttribute('tm'))
                          const valueNode = children(child(node, 'val'))[0]
                          const rawValue = valueNode && ['fltVal', 'strVal', 'intVal'].includes(valueNode.localName) ? valueNode.getAttribute('val') : null
                          const value = readCoordinateValue(rawValue, offset ?? 0)
                          const formula = node.getAttribute('fmla')?.trim() || undefined
                          return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && (hasNumericFormula || value >= -100 && value <= 100) && (formula === undefined || Boolean(formula && formula.length <= 16_384))
                              ? {offset, value, formula}
                              : undefined
                      }).filter((frame): frame is {offset: number; value: number; formula: string | undefined} => Boolean(frame))
                    : undefined
                const validCoordinateKeyframes = Boolean(coordinateKeyframes?.length === keyframeNodes.length && coordinateKeyframes.length >= 2 && coordinateKeyframes.every((frame, index) => index === 0 || frame.offset > coordinateKeyframes[index - 1]!.offset))
                const rawFrom = animation.getAttribute('from')
                const rawTo = animation.getAttribute('to')
                const rawBy = animation.getAttribute('by')
                // MS-OI29500 §19.5.1(h): numeric to-only/by-only tracks start at 0, then apply the target's saved geometry as an offset.
                const from = coordinateKeyframes?.[0]?.value ?? (rawFrom === null ? 0 : readCoordinateValue(rawFrom, 0))
                const by = rawBy === null || rawTo !== null ? undefined : readCoordinateValue(rawBy, 1)
                const to = coordinateKeyframes?.at(-1)?.value ?? (rawTo !== null ? readCoordinateValue(rawTo, 1) : by !== undefined && from !== undefined ? from + by : undefined)
                const hasFormulaKeyframes = Boolean(coordinateKeyframes?.some(frame => frame.formula !== undefined))
                const validCombination = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const validValues = tavList
                    ? validCoordinateKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : validCombination && from !== undefined && to !== undefined
                const behavior = child(animation, 'cBhvr')
                const additive = behavior?.getAttribute('additive')
                const accumulate = behavior?.getAttribute('accumulate')
                if (additive !== null && additive !== undefined && !['base', 'sum'].includes(additive) || accumulate !== null && accumulate !== undefined && !['none', 'always'].includes(accumulate)) {
                    warnings.add(`A ${property} animation with unsupported additive or cumulative composition was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                if (!validValues || from === undefined || to === undefined || (!hasFormulaKeyframes && (from < -100 || from > 100 || to < -100 || to > 100))) {
                    warnings.add(`A ${property} animation with unsupported values, formulas, or interpolation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Position')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const samples = sampleNumericAnimationFormula(
                    coordinateKeyframes || [{offset: 0, value: from}, {offset: 1, value: to}],
                    durationMs,
                    property,
                    element,
                    slideWidth,
                    slideHeight,
                    calculationMode as 'lin' | 'discrete'
                )
                if (!samples) {
                    warnings.add(`A ${property} animation formula used an unsupported variable or produced an out-of-range value; animation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:motionCoordinate:${property}:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'motion',
                    direction: 'in',
                    motionCoordinateProperty: property,
                    motionCoordinateKeyframeMode: calculationMode as 'lin' | 'discrete',
                    motionAdditive: additive === 'sum',
                    motionAccumulate: accumulate === 'always',
                    motionPathSamples: samples.map(frame => ({
                        progress: frame.offset,
                        x: property === 'ppt_x' ? (frame.value - baseline) * slideWidth : 0,
                        y: property === 'ppt_y' ? (frame.value - baseline) * slideHeight : 0
                    })),
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isOnOffProperty) {
                const propertyName = property as AnimationOnOffProperty
                if (!hasOnOffBaseline(element, propertyName)) {
                    warnings.add(`A ${propertyName} animation without a supported static baseline was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const onOffKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseOnOffValue(valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: boolean} => Boolean(frame))
                    : undefined
                const validOnOffKeyframes = Boolean(onOffKeyframes?.length === keyframeNodes.length && onOffKeyframes.every((frame, index) => index === 0 || frame.offset > onOffKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseOnOffValue(rawFrom)
                const to = rawTo === null ? undefined : parseOnOffValue(rawTo)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validOnOffKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                if (!validValues) {
                    warnings.add(`A ${propertyName} animation with unsupported string values or interpolation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Style on/off')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:onOff:${propertyName}:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'onOff',
                    direction: 'in',
                    onOffProperty: propertyName,
                    onOffFrom: from,
                    onOffTo: to,
                    onOffKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isStrokeDashProperty) {
                if (!hasOnOffBaseline(element, 'stroke.on')) {
                    warnings.add('A stroke.dashstyle animation without a supported static stroke baseline was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const strokeDashKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseStrokeDashStyle(valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: AnimatedStrokeDashStyle} => Boolean(frame))
                    : undefined
                const validStrokeDashKeyframes = Boolean(strokeDashKeyframes?.length === keyframeNodes.length && strokeDashKeyframes.every((frame, index) => index === 0 || frame.offset > strokeDashKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseStrokeDashStyle(rawFrom)
                const to = rawTo === null ? undefined : parseStrokeDashStyle(rawTo)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validStrokeDashKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                if (!validValues) {
                    warnings.add('A stroke.dashstyle animation with unsupported string values or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Stroke dash style')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = animationTarget.targetId + ':strokeDash:' + (timeNode?.getAttribute('id') || actions.size)
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'strokeDash',
                    direction: 'in',
                    // MS-OI29500 defines an omitted string start for to_only as empty;
                    // VML resolves an unspecified dashstyle to Solid.
                    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                    // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-dashstyle-attribute
                    strokeDashFrom: from ?? 'solid',
                    strokeDashTo: to,
                    strokeDashKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isLineArrowTypeProperty) {
                const lineArrowKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseLineArrow(valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: AnimatedLineArrow} => Boolean(frame))
                    : undefined
                const validLineArrowKeyframes = Boolean(lineArrowKeyframes?.length === keyframeNodes.length && lineArrowKeyframes.every((frame, index) => index === 0 || frame.offset > lineArrowKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseLineArrow(rawFrom)
                const to = rawTo === null ? undefined : parseLineArrow(rawTo)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validLineArrowKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                if (!validValues) {
                    warnings.add('A line arrow animation with unsupported string values or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                if (from === 'chevron' || to === 'chevron' || lineArrowKeyframes?.some(frame => frame.value === 'chevron')) warnings.add('The stroke arrow chevron preset is approximated with the open SVG marker.')
                const timing = animationTiming(timeNode, 'Line arrow')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                // MS-OI29500 starts to_only string animations at an empty string;
                // VML defines the default arrowhead as None.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vector-markup-language-object-model-reference
                const lineArrowFrom = from ?? (lineArrowKeyframes?.length ? undefined : 'none')
                const arrowProperty = property as AnimatedLineArrowProperty
                const key = animationTarget.targetId + ':lineArrow:' + arrowProperty + ':' + (timeNode?.getAttribute('id') || actions.size)
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'lineArrow',
                    direction: 'in',
                    lineArrowProperty: arrowProperty,
                    lineArrowFrom,
                    lineArrowTo: to,
                    lineArrowKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isLineArrowSizeProperty) {
                const lineArrowSizeProperty = property as AnimatedLineArrowSizeProperty
                const lineArrowSizeKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseLineArrowSize(lineArrowSizeProperty, valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: AnimatedLineArrowSize} => Boolean(frame))
                    : undefined
                const validLineArrowSizeKeyframes = Boolean(lineArrowSizeKeyframes?.length === keyframeNodes.length && lineArrowSizeKeyframes.every((frame, index) => index === 0 || frame.offset > lineArrowSizeKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseLineArrowSize(lineArrowSizeProperty, rawFrom)
                const to = rawTo === null ? undefined : parseLineArrowSize(lineArrowSizeProperty, rawTo)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validLineArrowSizeKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                if (!validValues) {
                    warnings.add('A line arrow size animation with unsupported string values or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Line arrow size')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                // MS-OI29500 starts to_only string animations at an empty string;
                // VML arrow width and length both default to Medium.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vector-markup-language-object-model-reference
                const lineArrowSizeFrom = from ?? (lineArrowSizeKeyframes?.length ? undefined : 'medium')
                const key = animationTarget.targetId + ':lineArrowSize:' + lineArrowSizeProperty + ':' + (timeNode?.getAttribute('id') || actions.size)
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'lineArrowSize',
                    direction: 'in',
                    lineArrowSizeProperty,
                    lineArrowSizeFrom,
                    lineArrowSizeTo: to,
                    lineArrowSizeKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (textStyleProperty) {
                const textStyleKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseTextStyleValue(textStyleProperty, valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: TextStyleAnimationValue} => Boolean(frame))
                    : undefined
                const validTextStyleKeyframes = Boolean(textStyleKeyframes?.length === keyframeNodes.length && textStyleKeyframes.every((frame, index) => index === 0 || frame.offset > textStyleKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseTextStyleValue(textStyleProperty, rawFrom)
                const to = rawTo === null ? undefined : parseTextStyleValue(textStyleProperty, rawTo)
                // MS-OI29500 starts to_only string animations at an empty value; use each VML text property's default.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-text-decoration-attribute
                // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-font-weight-attribute
                // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-font-style-attribute
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/ed0509b6-0550-4333-8266-884ac10bb5ba
                const omittedTextStyleStart: TextStyleAnimationValue | undefined = !textStyleKeyframes?.length
                    ? textStyleProperty === 'fontWeight' ? 'normal'
                        : textStyleProperty === 'fontStyle' || textStyleProperty === 'underline' || textStyleProperty === 'lineThrough' ? false
                            : textStyleProperty === 'textShadow' ? 'none'
                                : undefined
                    : undefined
                const textStyleFrom = from ?? omittedTextStyleStart
                const requestsMissingOutline = textStyleProperty === 'outline'
                    && (tavList ? textStyleKeyframes?.some(frame => frame.value === true) : from === true || to === true)
                    && !hasCompleteTextOutline(element)
                const unsupportedTextShadow = textStyleProperty === 'textShadow' && !hasCompleteTextShadow(element)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validTextStyleKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                    && !requestsMissingOutline && !unsupportedTextShadow
                if (!validValues) {
                    warnings.add(unsupportedTextShadow
                        ? 'A text shadow animation without complete static shadows was skipped.'
                        : requestsMissingOutline
                        ? 'A text outline animation without complete static outlines was skipped.'
                        : 'A generic text style animation with unsupported values, formulas, or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                if (textStyleProperty === 'emboss') warnings.add('A text emboss animation is approximated with CSS drop shadows; lighting and bevel depth may differ from Office.')
                const timing = animationTiming(timeNode, 'Text style')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:textStyle:${textStyleProperty}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'textStyle',
                    direction: 'in',
                    textStyleProperty,
                    textStyleFrom,
                    textStyleTo: to,
                    textStyleKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isVisibilityProperty) {
                // Microsoft defines visible/hidden as the string presets for style.visibility.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7b427ccc-3a1f-418b-821f-d78f7aed51c5
                const parseVisibility = (raw: string | null): AnimatedVisibility | undefined => {
                    const value = raw?.trim().toLowerCase()
                    return value === 'visible' || value === 'hidden' ? value : undefined
                }
                const visibilityKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseVisibility(valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: AnimatedVisibility} => Boolean(frame))
                    : undefined
                const validVisibilityKeyframes = Boolean(visibilityKeyframes?.length === keyframeNodes.length && visibilityKeyframes.every((frame, index) => index === 0 || frame.offset > visibilityKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseVisibility(rawFrom)
                const to = rawTo === null ? undefined : parseVisibility(rawTo)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validVisibilityKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                if (!validValues) {
                    warnings.add('A style.visibility animation with unsupported values or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Visibility')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:visibility:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
                // MS-OI29500 gives to_only an empty string start; VML resolves absent visibility to inherit.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-visibility-attribute
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'visibility',
                    direction: 'in',
                    visibilityFrom: from ?? 'inherit',
                    visibilityTo: to,
                    visibilityKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isGrayscaleProperty) {
                // Microsoft defines imageData.grayscale as a discrete false/true string property.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7b427ccc-3a1f-418b-821f-d78f7aed51c5
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa
                const parseGrayscale = (raw: string | null): boolean | undefined => {
                    const value = raw?.trim().toLowerCase()
                    return value === 'true' ? true : value === 'false' ? false : undefined
                }
                const grayscaleKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const value = valueNode?.localName === 'strVal' ? parseGrayscale(valueNode.getAttribute('val')) : undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && !node.hasAttribute('fmla')
                            ? {offset, value}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: boolean} => Boolean(frame))
                    : undefined
                const validGrayscaleKeyframes = Boolean(grayscaleKeyframes?.length === keyframeNodes.length && grayscaleKeyframes.every((frame, index) => index === 0 || frame.offset > grayscaleKeyframes[index - 1]!.offset))
                const from = rawFrom === null ? undefined : parseGrayscale(rawFrom)
                const to = rawTo === null ? undefined : parseGrayscale(rawTo)
                const validValues = calculationMode === 'discrete' && (tavList
                    ? validGrayscaleKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : to !== undefined && (rawFrom === null || from !== undefined))
                if (!validValues) {
                    warnings.add('An imageData.grayscale animation with unsupported values, formulas, or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Image grayscale')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:grayscale:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'grayscale',
                    direction: 'in',
                    // MS-OI29500 §19.5.1(h) makes an omitted string start empty; for grayscale that maps to false.
                    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                    // Microsoft lists false/true as the imageData.grayscale values.
                    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/0b35ab0b-2029-4180-8a76-8d784dac1d2b
                    grayscaleFrom: tavList ? from : from ?? false,
                    grayscaleTo: to,
                    grayscaleKeyframes,
                    durationMs: animationDurationMs(timeNode, timing),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (imageCropProperty) {
                // MS-OE376 lists all four imageData.crop* attributes as animation values.
                // Office stores them as ST_Fraction values; this renderer supports visible positive crop fractions.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7e506612-7a40-4d4e-95f4-e1f36173fe14
                const parseCropNumber = (raw: string | null): number | undefined => {
                    const value = raw?.trim()
                    if (!value) return undefined
                    const fixed = /^(-?\d+)f$/i.exec(value)
                    const parsed = fixed ? Number(fixed[1]) / 65_536 : parseOfficeAnimationNumber(value)
                    return parsed !== undefined && Number.isFinite(parsed) ? parsed : undefined
                }
                const parseCrop = (raw: string | null): number | undefined => {
                    const value = parseCropNumber(raw)
                    return value !== undefined && value >= 0 && value < 1 ? value : undefined
                }
                const cropKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const rawValue = valueNode && ['fltVal', 'strVal', 'intVal'].includes(valueNode.localName) ? valueNode.getAttribute('val') : null
                        const value = parseCrop(rawValue)
                        const formula = node.getAttribute('fmla')?.trim() || undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && (formula === undefined || formula.length <= 16_384)
                            ? {offset, value, formula}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: number; formula: string | undefined} => Boolean(frame))
                    : undefined
                const validCropKeyframes = Boolean(cropKeyframes?.length === keyframeNodes.length && cropKeyframes.every((frame, index) => index === 0 || frame.offset > cropKeyframes[index - 1]!.offset))
                const hasFormulaKeyframes = Boolean(cropKeyframes?.some(frame => frame.formula !== undefined))
                // MS-OI29500 defines from/to, from/by, to-only, and by-only combinations for p:anim;
                // numeric to-only and by-only tracks start at zero, and by is ignored when to is present.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                const validDirectCombination = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const from = rawFrom === null ? 0 : parseCrop(rawFrom)
                const by = rawBy !== null && rawTo === null ? parseCropNumber(rawBy) : undefined
                const to = rawTo !== null ? parseCrop(rawTo) : by !== undefined && from !== undefined ? from + by : undefined
                const oppositeProperty = imageCropProperty === 'left' ? 'right'
                    : imageCropProperty === 'right' ? 'left'
                    : imageCropProperty === 'top' ? 'bottom' : 'top'
                const oppositeCrop = element.imageCrop?.[oppositeProperty] || 0
                const cropValues = cropKeyframes?.map(frame => frame.value)
                    || [from ?? element.imageCrop?.[imageCropProperty] ?? 0, ...(to === undefined ? [] : [to])]
                const validStructure = tavList
                    ? ['lin', 'discrete'].includes(calculationMode) && validCropKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : calculationMode === 'lin' && validDirectCombination && to !== undefined && to >= 0 && to < 1 && (rawFrom === null || from !== undefined) && (rawBy === null || rawTo !== null || by !== undefined)
                if (!validStructure) {
                    warnings.add('An imageData.crop animation with unsupported values, formulas, or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Image crop')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const formulaSamples = hasFormulaKeyframes && cropKeyframes
                    ? sampleNumericAnimationFormula(cropKeyframes, durationMs, property as 'imageData.cropLeft' | 'imageData.cropTop' | 'imageData.cropRight' | 'imageData.cropBottom', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                const resolvedCropKeyframes = formulaSamples || cropKeyframes?.map(({offset, value}) => ({offset, value}))
                const resolvedCropValues = resolvedCropKeyframes?.map(frame => frame.value) ?? cropValues
                const fitsStaticOppositeCrop = resolvedCropValues.every(value => value >= 0 && value < 1 && value + oppositeCrop < 1)
                if (hasFormulaKeyframes && !formulaSamples || !fitsStaticOppositeCrop) {
                    warnings.add('An imageData.crop animation with unsupported values, formulas, or interpolation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:imageCrop:${imageCropProperty}:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'imageCrop',
                    direction: 'in',
                    imageCropProperty,
                    imageCropFrom: formulaSamples?.[0]?.value ?? from,
                    imageCropTo: formulaSamples?.at(-1)?.value ?? to,
                    imageCropKeyframes: resolvedCropKeyframes,
                    imageCropKeyframeMode: calculationMode as 'lin' | 'discrete',
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (imageToneProperty) {
                // MS-OE376 lists gain, blacklevel (spelled blackleve there), and gamma as imageData animation attributes.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa
                // Office uses ST_Fraction for these attributes and bounds blacklevel to [-0.5, 0.5].
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7e506612-7a40-4d4e-95f4-e1f36173fe14
                // MS-OI29500 §19.5.1(h) also allows from/to, from/by, to-only, and by-only numeric animations; omitted starts are 0.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                const isValidToneValue = (value: number): boolean => Number.isFinite(value) && (imageToneProperty === 'blacklevel'
                    ? value >= -0.5 && value <= 0.5
                    : imageToneProperty === 'gamma' ? value >= 0 && value <= 32_767 : value >= -32_768 && value <= 32_767)
                const parseToneNumber = (raw: string | null): number | undefined => {
                    const value = raw?.trim()
                    if (!value) return undefined
                    const fixed = /^(-?\d+)f$/i.exec(value)
                    const parsed = fixed ? Number(fixed[1]) / 65_536 : parseOfficeAnimationNumber(value)
                    return parsed !== undefined && Number.isFinite(parsed) ? parsed : undefined
                }
                const parseToneValue = (raw: string | null): number | undefined => {
                    const parsed = parseToneNumber(raw)
                    return parsed !== undefined && isValidToneValue(parsed) ? parsed : undefined
                }
                const toneKeyframes = tavList && keyframeNodes.length > 0 && keyframeNodes.length <= 256
                    ? keyframeNodes.map((node) => {
                        const valueNode = children(child(node, 'val'))[0]
                        const offset = parseFixedPercentage(node.getAttribute('tm'))
                        const rawValue = valueNode && ['fltVal', 'strVal', 'intVal'].includes(valueNode.localName) ? valueNode.getAttribute('val') : null
                        const value = parseToneValue(rawValue)
                        const formula = node.getAttribute('fmla')?.trim() || undefined
                        return offset !== undefined && offset >= 0 && offset <= 1 && value !== undefined && (formula === undefined || formula.length <= 16_384)
                            ? {offset, value, formula}
                            : undefined
                    }).filter((frame): frame is {offset: number; value: number; formula: string | undefined} => Boolean(frame))
                    : undefined
                const validToneKeyframes = Boolean(toneKeyframes?.length === keyframeNodes.length && toneKeyframes.every((frame, index) => index === 0 || frame.offset > toneKeyframes[index - 1]!.offset))
                const hasFormulaKeyframes = Boolean(toneKeyframes?.some(frame => frame.formula !== undefined))
                const from = rawFrom === null ? undefined : parseToneValue(rawFrom)
                const to = rawTo === null ? undefined : parseToneValue(rawTo)
                const by = rawBy === null || rawTo !== null ? undefined : parseToneNumber(rawBy)
                const directFrom = from ?? 0
                const directTo = to ?? (rawBy === null ? undefined : by === undefined ? undefined : directFrom + by)
                const validDirectValues = hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                    && (rawFrom === null || from !== undefined)
                    && (rawTo === null || to !== undefined)
                    && (rawBy === null || rawTo !== null || by !== undefined)
                    && directTo !== undefined
                    && isValidToneValue(directTo)
                const validValues = tavList
                    ? ['lin', 'discrete'].includes(calculationMode) && validToneKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : calculationMode === 'lin' && validDirectValues
                if (!validValues) {
                    warnings.add(`An ${property} animation with unsupported values, formulas, or interpolation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Image tone')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const formulaSamples = hasFormulaKeyframes && toneKeyframes
                    ? sampleNumericAnimationFormula(toneKeyframes, durationMs, property as 'imageData.gain' | 'imageData.blacklevel' | 'imageData.blackleve' | 'imageData.gamma', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && !formulaSamples) {
                    warnings.add(`An ${property} formula animation used an unsupported variable or produced an out-of-range value; animation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const resolvedToneKeyframes = formulaSamples || toneKeyframes?.map(({offset, value}) => ({offset, value}))
                const key = `${animationTarget.targetId}:imageTone:${imageToneProperty}:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'imageTone',
                    direction: 'in',
                    imageToneProperty,
                    imageToneFrom: formulaSamples?.[0]?.value ?? (tavList ? from : directFrom),
                    imageToneTo: formulaSamples?.at(-1)?.value ?? (tavList ? to : directTo),
                    imageToneKeyframes: resolvedToneKeyframes,
                    imageToneKeyframeMode: calculationMode as 'lin' | 'discrete',
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            // The Open XML SDK lists stroke.weight as an animation property; VML defines unitless stroke widths in EMU.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.attributename?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-strokeweight-attribute
            if (property === 'stroke.weight') {
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                const strokeWidthKeyframes = hasFormulaKeyframes ? parsedKeyframes : keyframes
                const validStrokeWidthKeyframes = Boolean(validKeyframeTrack && (hasFormulaKeyframes || strokeWidthKeyframes?.every(frame => frame.value <= MAX_ANIMATED_STROKE_WIDTH_EMU)))
                const validValues = tavList
                    ? validStrokeWidthKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                // MS-OI29500 19.5.1(h): numeric to-only/by-only p:anim starts at zero.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                const from = parsedKeyframes?.[0]?.value ?? (rawFrom === null ? 0 : parseNumericValue(rawFrom))
                const to = parsedKeyframes?.at(-1)?.value ?? (rawTo !== null ? parseNumericValue(rawTo) : rawBy !== null ? from + parseNumericValue(rawBy) : undefined)
                if (!validValues || to === undefined || !Number.isFinite(from) || !Number.isFinite(to) || (!hasFormulaKeyframes && (from < 0 || from > MAX_ANIMATED_STROKE_WIDTH_EMU || to < 0 || to > MAX_ANIMATED_STROKE_WIDTH_EMU))) {
                    warnings.add(hasFormulaKeyframes
                        ? 'A stroke.weight formula animation with unsupported interpolation or numeric keyframes was skipped.'
                        : 'A stroke.weight animation with missing or out-of-range numeric values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Stroke weight')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const strokeWidthFormulaSamples = hasFormulaKeyframes && parsedKeyframes
                    ? sampleNumericAnimationFormula(parsedKeyframes, durationMs, 'stroke.weight', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && !strokeWidthFormulaSamples) {
                    warnings.add('A stroke.weight formula used an unsupported variable or produced an out-of-range value; animation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                actions.set(`${animationTarget.targetId}:strokeWidth:${timeNode?.getAttribute('id') || actions.size}`, {
                    targetId: animationTarget.targetId,
                    effect: 'strokeWidth',
                    direction: 'in',
                    strokeWidthFrom: strokeWidthFormulaSamples?.[0]?.value ?? from,
                    strokeWidthTo: strokeWidthFormulaSamples?.at(-1)?.value ?? to,
                    strokeWidthKeyframes: keyframes,
                    strokeWidthKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['strokeWidthKeyframeMode']>) : undefined,
                    strokeWidthFormulaSamples,
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (isShadowOffsetProperty) {
                const shadowOffsetProperty = property as AnimationShadowOffsetProperty
                const hasTextShadowTarget = hasTextRange || ['lt', 'wd'].includes(iterate?.getAttribute('type') || '')
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                if (hasTextShadowTarget && hasFormulaKeyframes) {
                    warnings.add('A text-range shadow offset formula animation was skipped because formula variables cannot resolve per-run shadow baselines.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const shadowOffsetKeyframes = hasFormulaKeyframes ? parsedKeyframes : keyframes
                const validShadowOffsetKeyframes = Boolean(validKeyframeTrack && shadowOffsetKeyframes?.every(frame => Math.abs(frame.value) <= MAX_ANIMATED_SHADOW_OFFSET_EMU))
                const additive = behavior?.getAttribute('additive')
                const validComposition = additive === null || additive === undefined || ['base', 'repl', 'none'].includes(additive)
                const validValues = validComposition && (tavList
                    ? validShadowOffsetKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null))
                // PowerPoint numeric to-only/by-only tracks start at zero, independent of the saved property value.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e9355b93-5184-4ffe-8428-63a08858da2f
                const from = parsedKeyframes?.[0]?.value ?? (rawFrom === null ? 0 : parseNumericValue(rawFrom))
                const by = rawBy === null || rawTo !== null ? undefined : parseNumericValue(rawBy)
                const to = parsedKeyframes?.at(-1)?.value ?? (rawTo !== null ? parseNumericValue(rawTo) : by !== undefined ? from + by : undefined)
                if (!validValues || to === undefined || !Number.isFinite(from) || !Number.isFinite(to) || Math.abs(from) > MAX_ANIMATED_SHADOW_OFFSET_EMU || Math.abs(to) > MAX_ANIMATED_SHADOW_OFFSET_EMU || (by !== undefined && (!Number.isFinite(by) || Math.abs(by) > MAX_ANIMATED_SHADOW_OFFSET_EMU))) {
                    warnings.add('A generic shadow offset animation with unsupported composition or out-of-range values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Shadow offset')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const shadowOffsetFormulaSamples = hasFormulaKeyframes && parsedKeyframes
                    ? sampleNumericAnimationFormula(parsedKeyframes, durationMs, shadowOffsetProperty, element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && !shadowOffsetFormulaSamples) {
                    warnings.add(`A generic ${property} formula used an unsupported variable or produced an out-of-range value; animation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                actions.set(`${animationTarget.targetId}:shadowOffset:${shadowOffsetProperty}:${timeNode?.getAttribute('id') || actions.size}`, {
                    targetId: animationTarget.targetId,
                    effect: 'shadowOffset',
                    direction: 'in',
                    shadowOffsetProperty,
                    shadowOffsetFrom: shadowOffsetFormulaSamples?.[0]?.value ?? shadowOffsetKeyframes?.[0]?.value ?? from,
                    shadowOffsetTo: shadowOffsetFormulaSamples?.at(-1)?.value ?? shadowOffsetKeyframes?.at(-1)?.value ?? to,
                    shadowOffsetKeyframes: keyframes,
                    shadowOffsetFormulaSamples,
                    shadowOffsetKeyframeMode: tavList ? calculationMode as 'lin' | 'discrete' : undefined,
                    shadowOffsetAccumulate: behavior?.getAttribute('accumulate') === 'always',
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (!['style.fontSize', 'style.fontWeight'].includes(property || '')) {
                const opacityAdditive = behavior?.getAttribute('additive')
                const isOpacityProperty = ['style.opacity', 'fill.opacity', 'stroke.opacity', 'shadow.opacity'].includes(property || '')
                if (isOpacityProperty && opacityAdditive !== null && opacityAdditive !== undefined && !['base', 'sum', 'repl', 'mult', 'none'].includes(opacityAdditive)) {
                    warnings.add('A generic opacity animation with an unsupported additive mode was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                const formulaProperty = ['style.opacity', 'fill.opacity', 'stroke.opacity', 'shadow.opacity'].includes(property || '')
                const validOpacityKeyframes = formulaProperty && hasFormulaKeyframes
                    ? validKeyframeTrack
                    : Boolean(validKeyframes && keyframes?.every(frame => frame.value <= 1))
                const validValues = tavList
                    ? validOpacityKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const from = (formulaProperty ? parsedKeyframes?.[0]?.value : keyframes?.[0]?.value) ?? (rawFrom === null ? 0 : parseNumericValue(rawFrom))
                const to = (formulaProperty ? parsedKeyframes?.at(-1)?.value : keyframes?.at(-1)?.value) ?? (rawTo !== null ? parseNumericValue(rawTo) : rawBy !== null ? from + parseNumericValue(rawBy) : undefined)
                if (!validValues || (!tavList && (to === undefined || !Number.isFinite(from) || !Number.isFinite(to) || from < 0 || from > 1 || to < 0 || to > 1))) {
                    warnings.add(hasFormulaKeyframes
                        ? 'A generic opacity formula animation with an unsupported property, interpolation, or numeric keyframe was skipped.'
                        : 'A generic opacity animation with missing or out-of-range values was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const timing = animationTiming(timeNode, 'Opacity')
                if (!timing) {
                    handledTextAnimations.add(animation)
                    continue
                }
                const durationMs = animationDurationMs(timeNode, timing)
                const opacityFormulaSamples = hasFormulaKeyframes && formulaProperty && parsedKeyframes
                    ? sampleNumericAnimationFormula(parsedKeyframes, durationMs, property as NonNullable<PptxAnimation['opacityProperty']>, element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && !opacityFormulaSamples) {
                    warnings.add(`A generic ${property} formula used an unsupported variable or produced an out-of-range value; animation was skipped.`)
                    handledTextAnimations.add(animation)
                    continue
                }
                const rangeKey = hasTextRange ? ':' + (animationTarget.paragraphRange?.start ?? '') + ':' + (animationTarget.paragraphRange?.end ?? '') + ':' + (animationTarget.characterRange?.start ?? '') + ':' + (animationTarget.characterRange?.end ?? '') : ''
                actions.set(animationTarget.targetId + ':opacity:' + property + rangeKey + ':' + (timeNode?.getAttribute('id') || actions.size), {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'opacity',
                    opacityProperty: property as PptxAnimation['opacityProperty'],
                    direction: 'in',
                    opacityFrom: opacityFormulaSamples?.[0]?.value ?? keyframes?.[0]?.value ?? from,
                    opacityTo: opacityFormulaSamples?.at(-1)?.value ?? keyframes?.at(-1)?.value ?? to,
                    opacityKeyframes: keyframes,
                    opacityFormulaSamples,
                    opacityKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['opacityKeyframeMode']>) : undefined,
                    opacityAdditive: isOpacityProperty && (opacityAdditive === 'sum' || opacityAdditive === 'mult') ? opacityAdditive : undefined,
                    opacityAccumulate: behavior?.getAttribute('accumulate') === 'always',
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            if (property === 'style.fontWeight') {
                const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
                const validFontWeightKeyframes = Boolean(validKeyframeTrack && (hasFormulaKeyframes || parsedKeyframes?.every(frame => frame.value >= 1 && frame.value <= 1000)))
                const validValues = tavList
                    ? validFontWeightKeyframes && rawFrom === null && rawTo === null && rawBy === null
                    : hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
                const from = parsedKeyframes?.[0]?.value ?? (rawFrom === null ? 0 : parseNumericValue(rawFrom))
                const by = rawBy === null || rawTo !== null ? undefined : parseNumericValue(rawBy)
                const to = parsedKeyframes?.at(-1)?.value ?? (rawTo !== null ? parseNumericValue(rawTo) : by !== undefined ? from + by : undefined)
                const minimumFrom = rawFrom === null ? 0 : 1
                if (!validValues || !Number.isFinite(from) || !Number.isFinite(to) || to === undefined || (!hasFormulaKeyframes && (from < minimumFrom || from > 1000 || to < 1 || to > 1000))) {
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
                const fontWeightBaselines = [...new Set(element.paragraphs.flatMap(paragraph => paragraph.runs.map(run => run.bold ? 700 : 400)))]
                const formulaUsesFontWeightBaseline = Boolean(parsedKeyframes?.some(frame => /#style\.fontweight\b/i.test(frame.formula || '')))
                const formulaUsesRandom = Boolean(parsedKeyframes?.some(frame => /\brand\s*\(/i.test(frame.formula || '')))
                const needsPerBaselineFormulaSamples = Boolean(hasFormulaKeyframes && parsedKeyframes && formulaUsesFontWeightBaseline && fontWeightBaselines.length > 1)
                const fontWeightFormulaSeed = needsPerBaselineFormulaSamples && formulaUsesRandom ? Math.floor(Math.random() * 0x1_0000_0000) : undefined
                const fontWeightFormulaSamplesByBaseline = needsPerBaselineFormulaSamples && parsedKeyframes
                    ? fontWeightBaselines.map(baseline => ({
                        baseline,
                        samples: sampleNumericAnimationFormula(parsedKeyframes, durationMs, 'style.fontWeight', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete', baseline, fontWeightFormulaSeed)
                    }))
                    : undefined
                const fontWeightFormulaSamples = hasFormulaKeyframes && parsedKeyframes
                    ? fontWeightFormulaSamplesByBaseline
                        ? fontWeightFormulaSamplesByBaseline[0]?.samples
                        : sampleNumericAnimationFormula(parsedKeyframes, durationMs, 'style.fontWeight', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                    : undefined
                if (hasFormulaKeyframes && (!fontWeightFormulaSamples || fontWeightFormulaSamplesByBaseline?.some(entry => !entry.samples))) {
                    warnings.add('A generic text font-weight formula used an unsupported variable or produced an out-of-range value; animation was skipped.')
                    handledTextAnimations.add(animation)
                    continue
                }
                const key = `${animationTarget.targetId}:fontWeight:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'fontWeight',
                    direction: 'in',
                    fontWeightFrom: fontWeightFormulaSamples?.[0]?.value ?? from,
                    fontWeightTo: fontWeightFormulaSamples?.at(-1)?.value ?? to,
                    fontWeightKeyframes: keyframes,
                    fontWeightFormulaSamples,
                    fontWeightFormulaSamplesByBaseline: fontWeightFormulaSamplesByBaseline?.map(entry => ({baseline: entry.baseline, samples: entry.samples!})),
                    fontWeightKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['fontWeightKeyframeMode']>) : undefined,
                    durationMs,
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                handledTextAnimations.add(animation)
                continue
            }
            const hasFormulaKeyframes = Boolean(parsedKeyframes?.some(frame => frame.formula !== undefined))
            const validFontSizeKeyframes = Boolean(validKeyframeTrack && (hasFormulaKeyframes || parsedKeyframes?.every(frame => frame.value >= 0)))
            const validValues = tavList ? validFontSizeKeyframes && rawFrom === null && rawTo === null && rawBy === null : hasValidAnimateValueCombination(rawFrom !== null, rawTo !== null, rawBy !== null)
            const from = parsedKeyframes?.[0]?.value ?? (rawFrom === null ? (isGrowWithColorFontSize ? 1 : 0) : parseNumericValue(rawFrom))
            const to = parsedKeyframes?.at(-1)?.value ?? (rawTo === null ? (rawBy === null ? undefined : from + parseNumericValue(rawBy)) : parseNumericValue(rawTo))
            if (!validValues || to === undefined || !Number.isFinite(from) || !Number.isFinite(to) || (!hasFormulaKeyframes && (from < 0 || to < 0 || from > 100 || to > 100))) {
                warnings.add(hasFormulaKeyframes
                    ? 'A generic text font-size formula animation with unsupported interpolation or invalid numeric keyframes was skipped.'
                    : 'A generic text animation with missing or out-of-range font sizes was skipped.')
                handledTextAnimations.add(animation)
                continue
            }
            const timing = animationTiming(timeNode, 'Text font size')
            if (!timing) {
                handledTextAnimations.add(animation)
                continue
            }
            const durationMs = animationDurationMs(timeNode, timing)
            const fontSizeFormulaSamples = hasFormulaKeyframes && parsedKeyframes
                ? sampleNumericAnimationFormula(parsedKeyframes, durationMs, 'style.fontSize', element, slideWidth, slideHeight, calculationMode as 'lin' | 'discrete')
                : undefined
            if (hasFormulaKeyframes && !fontSizeFormulaSamples) {
                warnings.add('A generic text font-size formula used an unsupported variable or produced an out-of-range value; animation was skipped.')
                handledTextAnimations.add(animation)
                continue
            }
            const key = `${animationTarget.targetId}:fontSize:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'fontSize',
                direction: 'in',
                fontSizeFrom: fontSizeFormulaSamples?.[0]?.value ?? from,
                fontSizeTo: fontSizeFormulaSamples?.at(-1)?.value ?? to,
                fontSizeKeyframes: keyframes,
                fontSizeFormulaSamples,
                fontSizeKeyframeMode: tavList ? (calculationMode as NonNullable<PptxAnimation['fontSizeKeyframeMode']>) : undefined,
                durationMs,
                delayMs: delayMs + (isGrowWithColorFontSize && isWpsGrowWithColor ? wpsGrowWithColorChildStyleDelayMs : 0) + resolvedStartDelay(timeNode),
                ...timing
            })
            handledTextAnimations.add(animation)
        }
        for (const animation of ownedDescendants('animMotion')) {
            const animationTarget = target(animation)
            const timeNode = firstDescendant(animation, 'cTn')
            const element = elementsById.get(animationTarget.targetId)
            const path = animation.getAttribute('path')?.trim() || ''
            const behavior = child(animation, 'cBhvr')
            const rawOrigin = animation.getAttribute('origin')
            // PowerPoint defaults origin to parent only when the attribute is omitted; an explicit empty value is invalid.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            const origin = rawOrigin === null ? 'parent' : trimXmlWhitespace(rawOrigin)
            const fromNode = child(animation, 'from')
            const toNode = child(animation, 'to')
            const byNode = child(animation, 'by')
            const hasChildPosition = Boolean(fromNode || toNode || byNode)
            const behaviorFrom = behavior?.getAttribute('from') ?? null
            const behaviorTo = behavior?.getAttribute('to') ?? null
            const behaviorBy = behavior?.getAttribute('by') ?? null
            const hasBehaviorPosition = behaviorFrom !== null || behaviorTo !== null || behaviorBy !== null
            const hasPosition = hasChildPosition || hasBehaviorPosition
            const timing = animationTiming(timeNode, 'Motion')
            const preserveMotionTiming = () => preserveTimingOnly(animationTarget.targetId, timeNode, timing)
            if (!animationTarget.targetId) {
                warnings.add('A motion animation with an unsupported target or text range was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (!element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                warnings.add('A motion animation with an unsupported target or text range was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (!path && !hasPosition) {
                // PowerPoint treats an absent or empty path as a no-op that still consumes the behavior duration.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (path && hasPosition) {
                warnings.add('A motion animation combining a path with from/to/by position values was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (hasChildPosition && hasBehaviorPosition) {
                warnings.add('A motion animation combining common-behavior from/to/by values with position children was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (origin !== 'layout' && origin !== 'parent') {
                warnings.add('Motion animations without a supported layout or parent path are not rendered yet.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (hasBehaviorPosition && origin !== 'parent') {
                warnings.add('Common-behavior motion values with a non-parent origin are not rendered yet.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            const additive = behavior?.getAttribute('additive')
            if (additive !== null && additive !== undefined && !['base', 'sum'].includes(additive)) {
                warnings.add('A motion animation with an unsupported additive mode was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            const motionAttributeNames = children(child(behavior, 'attrNameLst'), 'attrName')
                .slice(0, 2)
                .map(node => normalizeAnimationAttributeName(node.textContent)?.toLowerCase())
            // MS-OI29500 allows up to two animMotion attributes and uses ppt_x/ppt_y when either slot is omitted.
            // Keep the slots in source order; the format does not require the two attributes to be distinct.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            const motionCoordinateProperties = [motionAttributeNames[0] || 'ppt_x', motionAttributeNames[1] || 'ppt_y'] as ['ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h', 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h']
            if (motionAttributeNames.some(property => !property)
                || motionCoordinateProperties.some(property => !['ppt_x', 'ppt_y', 'ppt_w', 'ppt_h'].includes(property))) {
                warnings.add('A motion animation with an empty or unsupported property was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            const hasDimensionMotion = motionCoordinateProperties.some(property => property === 'ppt_w' || property === 'ppt_h')
            const iteratesText = ['lt', 'wd'].includes(iterate?.getAttribute('type') || '')
            if (hasDimensionMotion && (animationTarget.paragraphRange || animationTarget.characterRange || iteratesText)) {
                warnings.add('A width or height motion path targeting only part of a text shape was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            const remapsMotionAxes = motionCoordinateProperties[0] !== 'ppt_x' || motionCoordinateProperties[1] !== 'ppt_y'
            if (!timing) {
                handledMotions.add(animation)
                continue
            }
            const angleSource = animation.getAttribute('rAng')
            const angleText = trimXmlWhitespace(angleSource)
            const rotationAngle = angleText === undefined ? 0 : /^[+-]?\d+$/.test(angleText) ? Number(angleText) : Number.NaN
            const readCoordinate = (raw: string | null): number | undefined => {
                if (raw === null) return undefined
                // ST_Percentage can carry a percent value followed by its integer fallback (for example, 56.7%457200).
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.rotationcenter?view=openxml-3.0.1
                return percentageValue(raw)
            }
            const readPosition = (node: Element | undefined): [number, number] | undefined => {
                if (!node) return undefined
                const x = readCoordinate(node.getAttribute('x'))
                const y = readCoordinate(node.getAttribute('y'))
                return x !== undefined && y !== undefined && Number.isFinite(x) && Number.isFinite(y) ? [x, y] : undefined
            }
            const rotationCenter = readPosition(child(animation, 'rCtr'))
            // PowerPoint defaults rAng to zero; nonzero values need an explicit slide-space rCtr.
            // rCtr is expressed as slide percentages; origin controls the path anchor, not rCtr eligibility.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a94c52e0-eb04-42b7-8b57-966bbd1f6061
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/90c939f8-7004-49c2-bb38-a6458363754e
            if (!Number.isInteger(rotationAngle) || rotationAngle < -2_147_483_648 || rotationAngle > 2_147_483_647 || (rotationAngle !== 0 && (!rotationCenter || hasDimensionMotion))) {
                warnings.add('A motion animation with an invalid rotation angle or unsupported rotation center was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            const angleRadians = (rotationAngle * Math.PI) / (60_000 * 180)
            const angleCosine = Math.cos(angleRadians),
                angleSine = Math.sin(angleRadians)
            const motionGeometry = element ? animationGeometryVariables(element, slideWidth, slideHeight) : undefined
            const formulaVariables: Record<'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h', number> =
                motionGeometry || {ppt_x: Number.NaN, ppt_y: Number.NaN, ppt_w: Number.NaN, ppt_h: Number.NaN}
            const readBehaviorPoint = (raw: string | null): [number, number] | {formulas: [string, string]} | undefined => {
                const value = trimXmlWhitespace(raw)
                if (!value || value.length > 32_768) return undefined
                const isFormulaPoint = value.startsWith('(') && value.endsWith(')')
                const pair = isFormulaPoint ? value.slice(1, -1) : value
                const commaSeparators: number[] = []
                const spaceSeparators: number[] = []
                let depth = 0
                for (let index = 0; index < pair.length; index++) {
                    if (pair[index] === '(') depth++
                    else if (pair[index] === ')') depth--
                    else if (depth === 0 && pair[index] === ',') commaSeparators.push(index)
                    else if (depth === 0 && /\s/.test(pair[index] || '') && (index === 0 || !/\s/.test(pair[index - 1] || ''))) spaceSeparators.push(index)
                }
                const separators = commaSeparators.length ? commaSeparators : spaceSeparators
                for (const separator of separators) {
                    const xFormula = pair.slice(0, separator).trim()
                    const yFormula = pair.slice(separator + 1).trim()
                    if (!xFormula || !yFormula || xFormula.length > 16_384 || yFormula.length > 16_384) continue
                    const variables = {...formulaVariables, $: 0.5}
                    const x = parseOfficeAnimationNumber(xFormula) ?? evaluateMotionFormula(xFormula, variables)
                    const y = parseOfficeAnimationNumber(yFormula) ?? evaluateMotionFormula(yFormula, variables)
                    if (x === undefined || y === undefined || !Number.isFinite(x) || !Number.isFinite(y)) continue
                    if (xFormula.includes('$') || yFormula.includes('$')) {
                        dynamicFormula = true
                        return {formulas: [xFormula, yFormula]}
                    }
                    return [x, y]
                }
                return undefined
            }
            const motionCoordinateBaselines = motionCoordinateProperties.map(property => formulaVariables[property])
            const rotateMotionOffset = (x: number, y: number): [number, number] => {
                if (rotationAngle === 0 || !rotationCenter || !motionGeometry) return [x, y]
                const centerX = rotationCenter[0] * slideWidth,
                    centerY = rotationCenter[1] * slideHeight
                const shapeCenterX = motionGeometry.ppt_x * slideWidth,
                    shapeCenterY = motionGeometry.ppt_y * slideHeight
                const dx = shapeCenterX + x - centerX,
                    dy = shapeCenterY + y - centerY
                return [centerX + dx * angleCosine - dy * angleSine - shapeCenterX, centerY + dx * angleSine + dy * angleCosine - shapeCenterY]
            }
            if (path.length > 100_000) {
                warnings.add('A motion path longer than 100,000 characters was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            // pathEditMode is editor metadata in Office and does not alter playback coordinates.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            // Office motion path commands use slide-size fractions and absolute/relative coordinates.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            const tokens: Array<string | {coordinates: [number, number]} | {formulas: [string, string]}> = []
            const numberPattern = /[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/iy
            const parseOfficeMotionCoordinate = (value: string): number | undefined => {
                // Office drops the decimal point when a fractional coordinate omits its leading zero.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
                const officeValue = /^[+-]?\.\d+(?:e[-+]?\d+)?$/i.test(value) ? value.replace('.', '') : value
                const coordinate = Number(officeValue)
                return Number.isFinite(coordinate) ? coordinate : undefined
            }
            let scan = 0
            let invalidPathSyntax = false
            let tooManyTokens = false
            let unsupportedFormula = false
            let dynamicFormula = false
            if (hasPosition) {
                // MS-OI29500 §19.5.22 defines cBhvr@from/to/by motion values as 2D points,
                // including parenthesized formulas. These are relative to the target's current position.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/b9a20a78-6401-4a8b-8ca9-b31ead409132
                const hasFrom = hasBehaviorPosition ? behaviorFrom !== null : Boolean(fromNode)
                const hasTo = hasBehaviorPosition ? behaviorTo !== null : Boolean(toNode)
                const hasBy = hasBehaviorPosition ? behaviorBy !== null : Boolean(byNode)
                const positionsAreValid = hasValidAnimateValueCombination(hasFrom, hasTo, hasBy)
                type MotionPoint = [number, number] | {formulas: [string, string]}
                const parsePoint = (node: Element | undefined, raw: string | null): MotionPoint | undefined => hasBehaviorPosition
                    ? readBehaviorPoint(raw)
                    : readPosition(node)
                const from: MotionPoint | undefined = hasFrom
                    ? parsePoint(fromNode, behaviorFrom)
                    : [0, 0]
                const by: MotionPoint | undefined = hasBy && !hasTo ? parsePoint(byNode, behaviorBy) : undefined
                const addPoints = (left: MotionPoint, right: MotionPoint): MotionPoint => {
                    if (Array.isArray(left) && Array.isArray(right)) return [left[0] + right[0], left[1] + right[1]]
                    const leftValues = Array.isArray(left) ? left.map(String) as [string, string] : left.formulas
                    const rightValues = Array.isArray(right) ? right.map(String) as [string, string] : right.formulas
                    return {formulas: [`(${leftValues[0]})+(${rightValues[0]})`, `(${leftValues[1]})+(${rightValues[1]})`]}
                }
                // When both to and by are present, PowerPoint uses to and ignores by.
                const to: MotionPoint | undefined = hasTo
                    ? parsePoint(toNode, behaviorTo)
                    : by && from
                        ? addPoints(from, by)
                        : undefined
                const pointValues = (position: MotionPoint): [number, number] => Array.isArray(position)
                    ? position
                    : [
                          evaluateMotionFormula(position.formulas[0], {...formulaVariables, $: 0.5}) ?? Number.NaN,
                          evaluateMotionFormula(position.formulas[1], {...formulaVariables, $: 0.5}) ?? Number.NaN
                      ]
                const withinPowerPointRange = (position: MotionPoint | undefined) => Boolean(position
                    && pointValues(position).every(value => Number.isFinite(value) && value >= -2_147_483_625 && value <= 2_147_483_625))
                if (!positionsAreValid || !from || !to || !withinPowerPointRange(from) || !withinPowerPointRange(to)) {
                    invalidPathSyntax = true
                } else {
                    // from/to/by x/y values are slide-size percentage changes. Convert them into
                    // the path's coordinate space so the existing path sampler and renderer can
                    // preserve the same pacing, repeats, rotation, and additive behavior.
                    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/0ce2da92-a863-43f5-be08-8a600f8ba222
                    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/f8515cce-689b-42c6-aaa1-816d25c1fd7b
                    const toPathCoordinates = (position: MotionPoint): MotionPoint => {
                        if (!hasBehaviorPosition && origin === 'layout') return position
                        if (Array.isArray(position)) return [position[0] + motionCoordinateBaselines[0]!, position[1] + motionCoordinateBaselines[1]!]
                        return {formulas: [
                            `(${position.formulas[0]})+(${motionCoordinateBaselines[0]})`,
                            `(${position.formulas[1]})+(${motionCoordinateBaselines[1]})`
                        ]}
                    }
                    const pathPoint = (position: MotionPoint): {coordinates: [number, number]} | {formulas: [string, string]} =>
                        Array.isArray(position) ? {coordinates: position} : {formulas: position.formulas}
                    tokens.push('M', pathPoint(toPathCoordinates(from)), 'L', pathPoint(toPathCoordinates(to)), 'E')
                }
            }
            while (!hasPosition && scan < path.length && !invalidPathSyntax) {
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
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (invalidPathSyntax || !tokens.length || !element) {
                if (hasPosition) warnings.add('A motion animation with an invalid from/to/by coordinate combination was skipped.')
                else if (unsupportedFormula) warnings.add('A motion path with an unsupported or invalid coordinate formula was skipped.')
                else warnings.add('An invalid or unsupported motion path was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            if (!Number.isFinite(formulaVariables.ppt_x) || !Number.isFinite(formulaVariables.ppt_y) || !Number.isFinite(formulaVariables.ppt_w) || !Number.isFinite(formulaVariables.ppt_h)) {
                warnings.add('An invalid or unsupported motion path was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            // Anchor motion points to the target's slide-space center. Group transforms are already
            // baked into the target geometry, so applying the group matrix here would scale offsets twice.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
            const point = (x: number, y: number): [number, number] => {
                const relativeLayoutPath = origin === 'layout'
                return [
                    (relativeLayoutPath ? x : x - motionCoordinateBaselines[0]!) * slideWidth,
                    (relativeLayoutPath ? y : y - motionCoordinateBaselines[1]!) * slideHeight
                ]
            }
            const initialPathPosition: [number, number] = origin === 'layout'
                ? [0, 0]
                : [motionCoordinateBaselines[0]!, motionCoordinateBaselines[1]!]
            const limit = 100 * Math.max(slideWidth, slideHeight)
            const buildPath = (progress: number) => {
                const pathPoints: NonNullable<PptxAnimation['motionPath']> = []
                const pathControls: [number, number][] = []
                let cursor = 0
                // Office defines L and C from the current point; when a path omits M,
                // begin at the target's existing position instead of rejecting the path.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/498c3cfa-652c-49b3-a82c-33fd94468af8
                let current: [number, number] = [...initialPathPosition]
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
                if (tokens[0] !== 'M' && tokens[0] !== 'm') append([...initialPathPosition], true)
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
                    if (typeof y !== 'string') return undefined
                    const coordinateX = parseOfficeMotionCoordinate(x)
                    const coordinateY = parseOfficeMotionCoordinate(y)
                    return coordinateX === undefined || coordinateY === undefined ? undefined : [coordinateX, coordinateY]
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
                        const previous = current
                        current = action === 'm' ? [previous[0] + coordinates[0], previous[1] + coordinates[1]] : coordinates
                        // Each moveto begins a new subpath; a following Z closes to this point.
                        pathStart = [...current]
                        // M snaps to a point; disconnected subpaths add no travel distance.
                        append(current, true)
                        continue
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
                        const curveTolerance = Math.max(slideWidth, slideHeight) / 10_000
                        const appendCurve = (p0: [number, number], p1: [number, number], p2: [number, number], p3: [number, number], depth = 0) => {
                            const chord = Math.hypot(p3[0] - p0[0], p3[1] - p0[1])
                            const controlLength = Math.hypot(p1[0] - p0[0], p1[1] - p0[1])
                                + Math.hypot(p2[0] - p1[0], p2[1] - p1[1])
                                + Math.hypot(p3[0] - p2[0], p3[1] - p2[1])
                            const controlDeviation = chord > 0
                                ? Math.max(
                                    Math.abs((p3[0] - p0[0]) * (p0[1] - p1[1]) - (p0[0] - p1[0]) * (p3[1] - p0[1])) / chord,
                                    Math.abs((p3[0] - p0[0]) * (p0[1] - p2[1]) - (p0[0] - p2[0]) * (p3[1] - p0[1])) / chord
                                )
                                : Math.max(Math.hypot(p1[0] - p0[0], p1[1] - p0[1]), Math.hypot(p2[0] - p0[0], p2[1] - p0[1]))
                            if (depth >= 7 || (controlDeviation <= curveTolerance && controlLength - chord <= curveTolerance)) {
                                if (depth >= 7 && (controlDeviation > curveTolerance || controlLength - chord > curveTolerance)) {
                                    warnings.add('A high-curvature motion path reached its adaptive sampling limit.')
                                }
                                const previous = pathPoints.at(-1)
                                if (previous) distance += Math.hypot(p3[0] - previous.x, p3[1] - previous.y)
                                pathPoints.push({x: p3[0], y: p3[1], distance})
                                return
                            }
                            const midpoint = (a: [number, number], b: [number, number]): [number, number] => [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2]
                            const p01 = midpoint(p0, p1), p12 = midpoint(p1, p2), p23 = midpoint(p2, p3)
                            const p012 = midpoint(p01, p12), p123 = midpoint(p12, p23), split = midpoint(p012, p123)
                            appendCurve(p0, p01, p012, split, depth + 1)
                            appendCurve(split, p123, p23, p3, depth + 1)
                        }
                        appendCurve(from, first, second, to)
                        current = end
                        drawingSegments++
                    } else if (action.toLowerCase() === 'z') {
                        const loopTarget = pathStart || initialPathPosition
                        append(loopTarget)
                        current = [...loopTarget]
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
                    preserveMotionTiming()
                    handledMotions.add(animation)
                    continue
                }
            }
            if (pathGeometry.invalid) {
                warnings.add('A motion animation with out-of-range coordinates was skipped.')
                preserveMotionTiming()
                handledMotions.add(animation)
                continue
            }
            const key = `${animationTarget.targetId}:motion:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}:${timeNode?.getAttribute('id') || actions.size}`
            actions.set(key, {
                targetId: animationTarget.targetId,
                paragraphRange: animationTarget.paragraphRange,
                characterRange: animationTarget.characterRange,
                effect: 'motion',
                direction: 'in',
                motionPath: pathGeometry.points,
                motionPathLength: pathGeometry.length,
                motionPathSamples,
                ...(remapsMotionAxes ? {motionCoordinateProperties} : {}),
                motionAdditive: additive === 'sum',
                motionAccumulate: behavior?.getAttribute('accumulate') === 'always',
                durationMs,
                delayMs: delayMs + resolvedStartDelay(timeNode),
                ...timing
            })
            handledMotions.add(animation)
        }
        for (const set of ownedDescendants('set')) {
            const attribute = normalizeAnimationAttributeName(firstDescendant(set, 'attrName')?.textContent)
            const value = firstDescendant(set, 'strVal')?.getAttribute('val')
            const behavior = child(set, 'cBhvr')
            const animationTarget = target(behavior || set)
            const element = elementsById.get(animationTarget.targetId)
            const parseSetNumber = (raw: string | null | undefined) => parseOfficeAnimationNumber(raw) ?? Number.NaN
            // Microsoft defines p:set value handling in `set` and the accepted text-style presets in `to`.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7d6aef67-5dbc-456c-bcf5-f08011bbf5bf
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7b427ccc-3a1f-418b-821f-d78f7aed51c5
            if (attribute === 'style.fontStyle' || attribute === 'style.fontFamily' || attribute === 'style.textDecorationUnderline' || attribute === 'style.textDecorationLineThrough' || attribute === 'style.textTransform' || attribute === 'style.textEffectOutline' || attribute === 'style.textShadow' || attribute === 'style.textEffectEmboss') {
                if (!animationTarget.targetId || !element?.paragraphs.length || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported target was skipped.`)
                    continue
                }
                const property: TextStyleAnimationProperty = attribute === 'style.fontStyle' ? 'fontStyle' : attribute === 'style.fontFamily' ? 'fontFamily' : attribute === 'style.textDecorationUnderline' ? 'underline' : attribute === 'style.textDecorationLineThrough' ? 'lineThrough' : attribute === 'style.textTransform' ? 'textTransform' : attribute === 'style.textEffectOutline' ? 'outline' : attribute === 'style.textShadow' ? 'textShadow' : 'emboss'
                const textStyleTo = parseTextStyleValue(property, child(child(set, 'to'), 'strVal')?.getAttribute('val') ?? null)
                if (textStyleTo === undefined) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported value was skipped.`)
                    continue
                }
                if (property === 'outline' && textStyleTo && !hasCompleteTextOutline(element)) {
                    warnings.add('A text outline p:set without complete static outlines was skipped.')
                    continue
                }
                if (property === 'textShadow' && !hasCompleteTextShadow(element)) {
                    warnings.add('A text shadow p:set without complete static shadows was skipped.')
                    continue
                }
                if (property === 'emboss') warnings.add('A text emboss animation is approximated with CSS drop shadows; lighting and bevel depth may differ from Office.')
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set text style')
                if (!timing) continue
                const key = `${animationTarget.targetId}:textStyle:set:${property}:${timeNode?.getAttribute('id') || actions.size}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'textStyle',
                    direction: 'in',
                    textStyleProperty: property,
                    textStyleFrom: textStyleTo,
                    textStyleTo,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'fill.type') {
                if (!animationTarget.targetId || element?.kind !== 'shape' || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange) {
                    warnings.add('A fill.type p:set animation with an unsupported target was skipped.')
                    continue
                }
                const fillType = child(child(set, 'to'), 'strVal')?.getAttribute('val')?.trim()
                const hasSolidBaseline = element.fill !== 'transparent' && !/(?:gradient\(|url\()/i.test(element.fill)
                const hasAnimatedFillColor = [...actions.values()].some(action => action.targetId === animationTarget.targetId && action.effect === 'color' && action.colorProperty === 'fillcolor')
                if (fillType !== 'solid' || (!hasSolidBaseline && !hasAnimatedFillColor)) {
                    warnings.add('A fill.type p:set animation without a supported solid color was skipped.')
                    continue
                }
                // A solid p:set needs no separate renderer state when its static or animated color is already a CSS solid color.
                solidFillTypeTargets.add(animationTarget.targetId)
                continue
            }
            if (attribute === 'fill.on' || attribute === 'stroke.on' || attribute === 'shadow.on') {
                const property = attribute as AnimationOnOffProperty
                const hasFillOnBaseline = property === 'fill.on' && element?.kind === 'shape'
                    && (hasOnOffBaseline(element, property) || solidFillTypeTargets.has(animationTarget.targetId))
                if (!animationTarget.targetId || !element || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element) || !(property === 'fill.on' ? hasFillOnBaseline : hasOnOffBaseline(element, property))) {
                    warnings.add(`A ${attribute} p:set animation without a supported whole-object baseline was skipped.`)
                    continue
                }
                const on = parseOnOffValue(child(child(set, 'to'), 'strVal')?.getAttribute('val') ?? null)
                if (on === undefined) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported value was skipped.`)
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set style on/off')
                if (!timing) continue
                const key = `${animationTarget.targetId}:onOff:${property}:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'onOff',
                    direction: 'in',
                    onOffProperty: property,
                    onOffFrom: on,
                    onOffTo: on,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'stroke.dashstyle') {
                if (!animationTarget.targetId || !element || !['shape', 'line'].includes(element.kind) || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element) || !hasOnOffBaseline(element, 'stroke.on')) {
                    warnings.add('A stroke.dashstyle p:set animation without a supported whole-object stroke baseline was skipped.')
                    continue
                }
                const strokeDash = parseStrokeDashStyle(child(child(set, 'to'), 'strVal')?.getAttribute('val') ?? null)
                if (!strokeDash) {
                    warnings.add('A stroke.dashstyle p:set animation with an unsupported value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set stroke dash style')
                if (!timing) continue
                const key = animationTarget.targetId + ':strokeDash:set:' + (timeNode?.getAttribute('id') || actions.size)
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'strokeDash',
                    direction: 'in',
                    strokeDashFrom: strokeDash,
                    strokeDashTo: strokeDash,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'stroke.startArrow' || attribute === 'stroke.endArrow') {
                if (!animationTarget.targetId || !element || element.kind !== 'line' || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element)) {
                    warnings.add('A line arrow p:set animation with an unsupported target was skipped.')
                    continue
                }
                const lineArrow = parseLineArrow(child(child(set, 'to'), 'strVal')?.getAttribute('val') ?? null)
                if (!lineArrow) {
                    warnings.add('A line arrow p:set animation with an unsupported value was skipped.')
                    continue
                }
                if (lineArrow === 'chevron') warnings.add('The stroke arrow chevron preset is approximated with the open SVG marker.')
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set line arrow')
                if (!timing) continue
                const arrowProperty = attribute as AnimatedLineArrowProperty
                const key = animationTarget.targetId + ':lineArrow:set:' + arrowProperty + ':' + (timeNode?.getAttribute('id') || actions.size)
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'lineArrow',
                    direction: 'in',
                    lineArrowProperty: arrowProperty,
                    lineArrowFrom: lineArrow,
                    lineArrowTo: lineArrow,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'stroke.startArrowWidth' || attribute === 'stroke.endArrowWidth' || attribute === 'stroke.startArrowLength' || attribute === 'stroke.endArrowLength') {
                if (!animationTarget.targetId || !element || element.kind !== 'line' || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element)) {
                    warnings.add('A line arrow size p:set animation with an unsupported target was skipped.')
                    continue
                }
                const property = attribute as AnimatedLineArrowSizeProperty
                const lineArrowSize = parseLineArrowSize(property, child(child(set, 'to'), 'strVal')?.getAttribute('val') ?? null)
                if (!lineArrowSize) {
                    warnings.add('A line arrow size p:set animation with an unsupported value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set line arrow size')
                if (!timing) continue
                const key = animationTarget.targetId + ':lineArrowSize:set:' + property + ':' + (timeNode?.getAttribute('id') || actions.size)
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'lineArrowSize',
                    direction: 'in',
                    lineArrowSizeProperty: property,
                    lineArrowSizeFrom: lineArrowSize,
                    lineArrowSizeTo: lineArrowSize,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'style.rotation') {
                if (!animationTarget.targetId || !element || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange) {
                    warnings.add('A style.rotation p:set animation with an unsupported target was skipped.')
                    continue
                }
                const to = child(set, 'to')
                const rawRotation = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const rotationText = rawRotation?.trim() || ''
                // WPS serializes this p:set value as degrees (for example, -45.0); generic p:anim rotation tracks use 60,000ths of a degree.
                const rotation = rotationText ? parseSetNumber(rotationText) : Number.NaN
                if (!Number.isFinite(rotation) || Math.abs(rotation) > 2_147_483_647 / 60_000) {
                    warnings.add('A style.rotation p:set animation with a missing or out-of-range value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set rotation')
                if (!timing) continue
                const key = `${animationTarget.targetId}:rotation:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'rotation',
                    direction: 'in',
                    rotationFrom: rotation,
                    rotationTo: rotation,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'style.fontWeight') {
                // MS-OE376 specifies the p:set string presets none, normal, and bold for style.fontWeight.
                // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7b427ccc-3a1f-418b-821f-d78f7aed51c5
                if (!animationTarget.targetId || !element?.paragraphs.length || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                    warnings.add('A style.fontWeight p:set animation with an unsupported target was skipped.')
                    continue
                }
                const fontWeightValue = value?.trim().toLowerCase()
                if (!['none', 'normal', 'bold'].includes(fontWeightValue || '')) {
                    warnings.add('A style.fontWeight p:set animation with an unsupported value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set font weight')
                if (!timing) continue
                const fontWeight = fontWeightValue === 'bold' ? 700 : 400
                const key = `${animationTarget.targetId}:fontWeight:set:${timeNode?.getAttribute('id') || actions.size}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'fontWeight',
                    direction: 'in',
                    fontWeightFrom: fontWeight,
                    fontWeightTo: fontWeight,
                    fontWeightReset: fontWeightValue === 'none',
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            // p:set holds one fixed style.fontSize factor; Microsoft's animation example uses 1.5 for 150%.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.setbehavior?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
            if (attribute === 'style.fontSize') {
                if (!animationTarget.targetId || !element?.paragraphs.length || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                    warnings.add('A style.fontSize p:set animation with an unsupported target was skipped.')
                    continue
                }
                const to = child(set, 'to')
                const rawFontSize = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const fontSize = rawFontSize?.trim() ? parseSetNumber(rawFontSize) : Number.NaN
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
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'stroke.weight') {
                if (!animationTarget.targetId || !element || !['shape', 'line'].includes(element.kind) || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange || !hasValidCharacterRange(animationTarget, element)) {
                    warnings.add('A stroke.weight p:set animation with an unsupported target was skipped.')
                    continue
                }
                const to = child(set, 'to')
                const rawValue = child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val') ?? child(to, 'strVal')?.getAttribute('val')
                const strokeWidth = rawValue?.trim() ? parseSetNumber(rawValue) : Number.NaN
                if (!Number.isFinite(strokeWidth) || strokeWidth < 0 || strokeWidth > MAX_ANIMATED_STROKE_WIDTH_EMU) {
                    warnings.add('A stroke.weight p:set animation with a missing or out-of-range value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, 'p:set stroke weight')
                if (!timing) continue
                const key = `${animationTarget.targetId}:strokeWidth:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'strokeWidth',
                    direction: 'in',
                    strokeWidthFrom: strokeWidth,
                    strokeWidthTo: strokeWidth,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'shadow.offset.x' || attribute === 'shadow.offset.y') {
                if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasShadow(element) || !hasValidTextRange(animationTarget, element)) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported target was skipped.`)
                    continue
                }
                const to = child(set, 'to')
                const rawValue = child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val') ?? child(to, 'strVal')?.getAttribute('val')
                const offset = rawValue?.trim() ? parseSetNumber(rawValue) : Number.NaN
                if (!Number.isFinite(offset) || Math.abs(offset) > MAX_ANIMATED_SHADOW_OFFSET_EMU) {
                    warnings.add(`A ${attribute} p:set animation with a missing or out-of-range value was skipped.`)
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const timing = animationTiming(timeNode, `p:set ${attribute}`)
                if (!timing) continue
                const key = `${animationTarget.targetId}:shadowOffset:${attribute}:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'shadowOffset',
                    direction: 'in',
                    shadowOffsetProperty: attribute,
                    shadowOffsetFrom: offset,
                    shadowOffsetTo: offset,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
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
                    ((animationTarget.paragraphRange || animationTarget.characterRange) && !['style.opacity', 'shadow.opacity'].includes(property)) ||
                    !hasValidTextRange(animationTarget, element) ||
                    (property === 'fill.opacity' && element?.kind !== 'shape') ||
                    (property === 'stroke.opacity' && (!element || !supportsStrokeOpacity(element))) ||
                    (property === 'shadow.opacity' && (!element || !hasShadow(element)))
                if (unsupportedTarget) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported target was skipped.`)
                    continue
                }
                const to = child(set, 'to')
                const rawValue = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const opacity = rawValue?.trim() ? parseSetNumber(rawValue) : Number.NaN
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
                    delayMs: delayMs + resolvedStartDelay(timeNode),
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
                    !hasValidTextRange(animationTarget, element) ||
                    ((animationTarget.paragraphRange || animationTarget.characterRange) && !['style.color', 'shadow.color'].includes(property)) ||
                    (property === 'style.color' && element.kind === 'table' && animationTarget.paragraphRange) ||
                    (property === 'fillcolor' && (element?.kind !== 'shape' || /gradient\(/i.test(element.fill))) ||
                    (property === 'stroke.color' && (!['shape', 'line'].includes(element?.kind || '') || /gradient\(/i.test(element.stroke))) ||
                    (property === 'shadow.color' && (!element || !hasShadow(element))) ||
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
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'color',
                    direction: 'in',
                    colorProperty: property,
                    colorFrom: color,
                    colorTo: color,
                    colorSpace: 'rgb',
                    colorDirection: 'cw',
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
                continue
            }
            if (attribute === 'imageData.grayscale') {
                if (!animationTarget.targetId || element?.kind !== 'picture' || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange) {
                    warnings.add('An imageData.grayscale p:set animation requires a whole-picture target.')
                    continue
                }
                const grayscaleValue = value?.trim().toLowerCase()
                if (grayscaleValue !== 'true' && grayscaleValue !== 'false') {
                    warnings.add('An imageData.grayscale p:set animation with an unsupported value was skipped.')
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const key = `${animationTarget.targetId}:grayscale:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'grayscale',
                    direction: 'in',
                    grayscaleTo: grayscaleValue === 'true',
                    durationMs: 0,
                    delayMs: delayMs + resolvedStartDelay(timeNode)
                })
                continue
            }
            const imageCropProperty = attribute === 'imageData.cropLeft' ? 'left'
                : attribute === 'imageData.cropTop' ? 'top'
                : attribute === 'imageData.cropRight' ? 'right'
                : attribute === 'imageData.cropBottom' ? 'bottom'
                : undefined
            if (imageCropProperty) {
                if (!animationTarget.targetId || element?.kind !== 'picture' || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange) {
                    warnings.add(`A ${attribute} p:set animation requires a whole-picture target.`)
                    continue
                }
                const to = child(set, 'to')
                const rawCrop = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const cropText = rawCrop?.trim() || ''
                const fixedCrop = /^(-?\d+)f$/i.exec(cropText)
                const cropValue = cropText ? fixedCrop ? Number(fixedCrop[1]) / 65_536 : parseSetNumber(cropText) : Number.NaN
                const oppositeProperty = imageCropProperty === 'left' ? 'right'
                    : imageCropProperty === 'right' ? 'left'
                    : imageCropProperty === 'top' ? 'bottom' : 'top'
                if (!Number.isFinite(cropValue) || cropValue < 0 || cropValue >= 1 || cropValue + (element.imageCrop?.[oppositeProperty] || 0) >= 1) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported crop value was skipped.`)
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const key = `${animationTarget.targetId}:imageCrop:${imageCropProperty}:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'imageCrop',
                    direction: 'in',
                    imageCropProperty,
                    imageCropFrom: cropValue,
                    imageCropTo: cropValue,
                    durationMs: 0,
                    delayMs: delayMs + resolvedStartDelay(timeNode)
                })
                continue
            }
            const imageToneProperty = attribute === 'imageData.gain' ? 'gain'
                : attribute === 'imageData.blacklevel' || attribute === 'imageData.blackleve' ? 'blacklevel'
                : attribute === 'imageData.gamma' ? 'gamma'
                : undefined
            if (imageToneProperty) {
                if (!animationTarget.targetId || element?.kind !== 'picture' || animationTarget.paragraphRange || animationTarget.characterRange || animationTarget.hasUnsupportedTextRange) {
                    warnings.add(`A ${attribute} p:set animation requires a whole-picture target.`)
                    continue
                }
                const to = child(set, 'to')
                const rawValue = child(to, 'strVal')?.getAttribute('val') ?? child(to, 'fltVal')?.getAttribute('val') ?? child(to, 'intVal')?.getAttribute('val')
                const text = rawValue?.trim() || ''
                const fixed = /^(-?\d+)f$/i.exec(text)
                const toneValue = text ? fixed ? Number(fixed[1]) / 65_536 : parseSetNumber(text) : Number.NaN
                const validToneValue = Number.isFinite(toneValue) && (imageToneProperty === 'blacklevel'
                    ? toneValue >= -0.5 && toneValue <= 0.5
                    : imageToneProperty === 'gamma' ? toneValue >= 0 && toneValue <= 32_767 : toneValue >= -32_768 && toneValue <= 32_767)
                if (!validToneValue) {
                    warnings.add(`A ${attribute} p:set animation with an unsupported value was skipped.`)
                    continue
                }
                const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
                const key = `${animationTarget.targetId}:imageTone:${imageToneProperty}:set:${timeNode?.getAttribute('id') || actions.size}`
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    effect: 'imageTone',
                    direction: 'in',
                    imageToneProperty,
                    imageToneFrom: toneValue,
                    imageToneTo: toneValue,
                    durationMs: 0,
                    delayMs: delayMs + resolvedStartDelay(timeNode)
                })
                continue
            }
            if (attribute !== 'style.visibility') {
                warnings.add(`A p:set animation for ${attribute || 'an unknown property'} is not rendered yet.`)
                continue
            }
            if (!animationTarget.targetId || !element || animationTarget.hasUnsupportedTextRange || !hasValidTextRange(animationTarget, element)) {
                warnings.add('A style.visibility p:set animation with an unsupported target was skipped.')
                continue
            }
            if (!['visible', 'hidden'].includes(value || '')) {
                warnings.add('A style.visibility p:set animation with an invalid value was skipped.')
                continue
            }
            const timeNode = child(behavior, 'cTn') || firstDescendant(set, 'cTn')
            const timing = animationTiming(timeNode, 'p:set visibility')
            if (!timing) continue
            const key = `${animationTarget.targetId}:visibility:set:${timeNode?.getAttribute('id') || actions.size}:${animationTarget.paragraphRange?.start ?? ''}:${animationTarget.paragraphRange?.end ?? ''}:${animationTarget.characterRange?.start ?? ''}:${animationTarget.characterRange?.end ?? ''}`
            if (!actions.has(key)) {
                actions.set(key, {
                    targetId: animationTarget.targetId,
                    paragraphRange: animationTarget.paragraphRange,
                    characterRange: animationTarget.characterRange,
                    effect: 'visibility',
                    direction: value === 'visible' ? 'in' : 'out',
                    // p:set applies its value at begin; cTn duration controls its active/fill lifetime.
                    effectProgress: 1,
                    visibilityFrom: value === 'visible' ? 'hidden' : 'visible',
                    visibilityTo: value === 'visible' ? 'visible' : 'hidden',
                    durationMs: animationDurationMs(timeNode, timing, 1),
                    delayMs: delayMs + resolvedStartDelay(timeNode),
                    ...timing
                })
            }
        }
        // tmPct staggers the containing effect, so an instantaneous visibility set
        // shares the same interval as its size/opacity tracks. Otherwise later
        // letters appear at full size before their Zoom animation starts.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timepercentage
        const parsedIterationDurationMs = [...actions.values()].reduce((duration, action) => Math.max(duration, Math.max(0, action.delayMs - delayMs) + action.durationMs), 0)
        const iterationDurationMs = isWpsGrowWithColor
            ? [...actions.values()].reduce((duration, action) => Math.max(duration, action.durationMs), 0)
            : parsedIterationDurationMs || (iterate
            ? ownedDescendants('cTn').reduce((duration, timeNode) => Math.max(duration, timelineTime(timeNode.getAttribute('dur'), 0)), 0)
            : 0)
        let groupedActions = [...actions.entries()].flatMap(([actionKey, action], actionIndex) => {
            const mediaActionKey = action.effect === 'media' ? `${sequenceKey}:${sequenceOrder}:${actionIndex}:${action.targetId}:${mediaTimeNodeIds.get(actionKey) ?? actionKey}` : undefined
            const mediaTimeNodeId = mediaTimeNodeIds.get(actionKey)
            if (mediaTimeNodeId && mediaActionKey) {
                if (sequenceKey === 'main') {
                    timeNodeSources.set(mediaTimeNodeId, {...(timeNodeSources.get(mediaTimeNodeId) || {}), mediaActionKey})
                }
            }
            const sequencedAction = {
                ...action,
                ...(mediaTimeNodeId ? {sourceTimeNodeId: mediaTimeNodeId} : {}),
                sequenceKey,
                sequenceOrder,
                fillSequenceScopes: [{key: sequenceKey, order: sequenceOrder}],
                mediaActionKey,
                mediaWaitForEndKeys: mediaWaitForEndKeys.length ? mediaWaitForEndKeys : undefined,
                mediaWaitDelayMs: mediaWaitForEndKeys.length ? mediaWaitDelayMs + action.delayMs - delayMs : undefined
            }
            const elementActions = elementIterationActions(sequencedAction)
            return elementActions ? elementActions.map(elementAction => ({...elementAction, sequenceKey, sequenceOrder})) : [{...sequencedAction, iteration: iterationRanges(sequencedAction)}]
        })
        const offsetEndConditionDurationMs = owner && child(owner, 'endCondLst')
            ? offsetEndDuration(owner, 'Animation group')
            : undefined
        const hasParentClockTransform = Boolean(owner && ['accel', 'decel', 'tmFilter', 'spd', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => owner.hasAttribute(attribute)))
        const groupEventTimeNodeSources = new Map(eventTimeNodeSources || [])
        const localTimeNodes = descendants(group, 'cTn')
        const localTimeNodesById = new Map(localTimeNodes.flatMap(timeNode => {
            const id = timeNode.getAttribute('id')
            return id ? [[id, timeNode] as const] : []
        }))
        if (!hasParentClockTransform) {
            const sourceActionsById = new Map<string, PptxAnimation[]>()
            for (const action of groupedActions) {
                const id = action.sourceTimeNodeId
                if (!id || action.effect === 'media' || id === owner?.getAttribute('id') || !localTimeNodesById.has(id)) continue
                const sourceActions = sourceActionsById.get(id) || []
                sourceActions.push(action)
                sourceActionsById.set(id, sourceActions)
            }
            for (const [id, sourceActions] of sourceActionsById) {
                const sourceNode = localTimeNodesById.get(id)!
                const startMs = Math.min(...sourceActions.map(action => action.delayMs))
                const naturalEndMs = Math.max(...sourceActions.map(action => action.delayMs + playbackDuration(action)))
                groupEventTimeNodeSources.set(id, {
                    startMs,
                    ...(!deferredEndConditionWarnings.has(sourceNode) && Number.isFinite(naturalEndMs) ? {endMs: naturalEndMs} : {})
                })
            }
            // Resolve child end conditions first, then make their resolved end events available to containers.
            for (let pass = 0; pass < localTimeNodes.length; pass++) {
                let resolvedAny = false
                for (const timeNode of localTimeNodes) {
                    const id = timeNode.getAttribute('id')
                    if (!id || timeNode === owner || !deferredEndConditionWarnings.has(timeNode)) continue
                    const relatedActions = groupedActions.filter(action => action.effect !== 'media' && action.sourceTimeNodeId === id)
                    if (!relatedActions.length) continue
                    const startMs = Math.min(...relatedActions.map(action => action.delayMs))
                    const durationMs = subTimeEventEndDuration(timeNode, startMs, groupEventTimeNodeSources)
                    if (durationMs === undefined) continue
                    const endMs = startMs + durationMs
                    deferredEndConditionWarnings.delete(timeNode)
                    parsedTimeNodeEndOffsets.set(timeNode, endMs)
                    groupedActions = groupedActions.map(action => action.sourceTimeNodeId === id
                        ? {...action, parentEndMs: Math.min(action.parentEndMs ?? Infinity, endMs)}
                        : action)
                    groupEventTimeNodeSources.set(id, {startMs, endMs})
                    resolvedAny = true
                }
                if (!resolvedAny) break
            }
        }
        const referencedEndConditionDurationMs = offsetEndConditionDurationMs === undefined
            && owner
            && !subTimeEntry
            && groupEventTimeNodeSources.size
            && deferredEndConditionWarnings.has(owner)
            ? subTimeEventEndDuration(owner, delayMs, groupEventTimeNodeSources)
            : undefined
        if (referencedEndConditionDurationMs !== undefined && owner) deferredEndConditionWarnings.delete(owner)
        const finiteEndConditionDurationMs = offsetEndConditionDurationMs ?? referencedEndConditionDurationMs
        let perIterationGroupTiming = false
        if (canApplyGroupTiming && owner) {
            const rawDuration = owner.getAttribute('dur')
            const declaredDurationMs = timelineTime(rawDuration, Number.NaN)
            const needsChildDerivedDuration = rawDuration === null || declaredDurationMs === 0
                || isIndefiniteTimelineTime(rawDuration) && Number.isFinite(finiteEndConditionDurationMs)
            // p:iterate repeats contained timing structures over letters/words/shapes; apply this parent curve to
            // each unit's single timeline while leaving its tmPct/tmAbs stagger on the wall clock.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.iterate?view=openxml-3.0.1
            const supportsPerIterationTiming = hasIteration && groupedActions.length > 0
                && groupedActions.every(action => Boolean(action.iteration?.ranges.length))
                && !['spd', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => owner.hasAttribute(attribute))
            perIterationGroupTiming = needsChildDerivedDuration && supportsPerIterationTiming
            const supportsSpeedOnlyIterationTiming = hasIteration && owner.hasAttribute('spd')
                && !['accel', 'decel', 'tmFilter', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => owner.hasAttribute(attribute))
            const canDeriveDuration = needsChildDerivedDuration && (!hasIteration || supportsPerIterationTiming || supportsSpeedOnlyIterationTiming) && groupedActions.length > 0 && groupedActions.every(action => {
                const childOffsetMs = action.delayMs - delayMs
                const childDurationMs = supportsPerIterationTiming ? basePlaybackDuration(action) : playbackDuration(action)
                return (action.effect !== 'media' || Number.isFinite(action.mediaDurationMs))
                    && Number.isFinite(childOffsetMs)
                    && childOffsetMs >= 0
                    && Number.isFinite(childDurationMs)
                    && childDurationMs >= 0
            })
            const childDerivedDurationMs = canDeriveDuration
                ? groupedActions.reduce((duration, action) => Math.max(duration, action.delayMs - delayMs + (supportsPerIterationTiming ? basePlaybackDuration(action) : playbackDuration(action))), 0)
                : undefined
            groupTiming = groupTimingFor(owner, 'Animation group', childDerivedDurationMs, finiteEndConditionDurationMs)
        }
        let timedActions = groupTiming
            ? applyGroupTiming(groupedActions, delayMs, groupTiming, perIterationGroupTiming, owner)
            : groupedActions
        if (groupEventTimeNodeSources.size) {
            const actionTimeNodes = localTimeNodes
            for (const timeNode of actionTimeNodes) {
                const timeNodeId = timeNode.getAttribute('id')
                if (!timeNodeId || !deferredEndConditionWarnings.has(timeNode)) continue
                const relatedActions = timedActions.filter(action => action.effect !== 'media' && action.sourceTimeNodeId === timeNodeId)
                if (!relatedActions.length) continue
                const timeNodeStartMs = Math.min(...relatedActions.map(action => action.delayMs))
                const endConditionDurationMs = subTimeEventEndDuration(timeNode, timeNodeStartMs, groupEventTimeNodeSources)
                if (endConditionDurationMs === undefined) continue
                const timeNodeEndMs = timeNodeStartMs + endConditionDurationMs
                deferredEndConditionWarnings.delete(timeNode)
                parsedTimeNodeEndOffsets.set(timeNode, timeNodeEndMs)
                timedActions = timedActions.map(action => action.sourceTimeNodeId === timeNodeId
                    ? {...action, parentEndMs: Math.min(action.parentEndMs ?? Infinity, timeNodeEndMs)}
                    : action)
            }
        }
        const rawDuration = owner?.getAttribute('dur')
        const declaredDurationMs = timelineTime(rawDuration, Number.NaN)
        let parentEndMs = groupTiming ? delayMs + groupTiming.playbackDurationMs : undefined
        if (parentEndMs === undefined && rawDuration && !isIndefiniteTimelineTime(rawDuration) && declaredDurationMs > 0) {
            parentEndMs = delayMs + declaredDurationMs
        } else if (parentEndMs === undefined && (rawDuration === null || declaredDurationMs === 0)) {
            const naturalEnd = timedActions.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), delayMs)
            if (Number.isFinite(naturalEnd)) parentEndMs = naturalEnd
        }
        if (Number.isFinite(finiteEndConditionDurationMs)) {
            parentEndMs = Math.min(parentEndMs ?? Infinity, delayMs + finiteEndConditionDurationMs!)
        }
        if (owner && Number.isFinite(parentEndMs)) parsedTimeNodeEndOffsets.set(owner, parentEndMs!)
        if (Number.isFinite(parentEndMs)) {
            return timedActions.map(action => ({
                ...action,
                parentEndMs: Math.min(action.parentEndMs ?? Infinity, parentEndMs!),
                ...parentFillEnd(action, owner, parentEndMs!)
            }))
        }
        return timedActions
    }

    const conditionalSequencePlans: Array<{step: PptxAnimation[]; actions: PptxAnimation[][]; concurrent: boolean; nextAction: 'none' | 'seek'; rewindsOnPrevious: boolean}> = []
    const appendSubTimeEntry = (
        entry: Element,
        ownerGroup: Element,
        parentTimeNode: Element,
        sequence: PptxAnimation[],
        sequenceKey: string,
        sequenceOrder: number,
        startMs: number,
        inheritedTimeNodeSources: Map<string, {startMs: number; endMs?: number}> = new Map()
    ) => {
        const hasImmediateSlideCondition = (sequenceNode: Element, listName: 'nextCondLst' | 'prevCondLst', eventName: 'onNext' | 'onPrev') => {
            const conditions = children(child(sequenceNode, listName), 'cond')
            // PowerPoint treats entries in nextCondLst/prevCondLst as alternatives:
            // any matching condition advances or rewinds the sequence.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a5b47f26-9eb6-4d45-acd6-e7183453307a
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/31d28bcd-cca2-42ea-95af-eb32b797e869
            const isSupported = (condition: Element) =>
                condition.getAttribute('evt') === eventName
                && (!condition.hasAttribute('delay') || (isFiniteTimelineTime(condition.getAttribute('delay')) && timelineTime(condition.getAttribute('delay'), Number.NaN) === 0))
                && condition.children.length === 1
                && condition.children[0]?.localName === 'tgtEl'
                && Boolean(firstDescendant(condition, 'sldTgt'))
            const hasSupportedAlternative = conditions.some(isSupported)
            if (hasSupportedAlternative && conditions.some(condition => !isSupported(condition))) {
                warnings.add(`Some ${listName} alternatives in a nested p:seq are unsupported; immediate slide-target navigation alternatives are applied.`)
            }
            return hasSupportedAlternative
        }
        const supportsConditionalSequence = (sequenceNode: Element, timeNode: Element, childCount: number, advancesOnNext: boolean) => {
            const concurrent = sequenceNode.getAttribute('concurrent')
            const nextAction = sequenceNode.getAttribute('nextAc')
            const previousAction = sequenceNode.getAttribute('prevAc')
            const duration = timeNode.getAttribute('dur')
            // Expanded child actions are separated by onNext boundaries, which are skipTimed rewind stops.
            return childCount > 1
                && [null, '0', 'false', '1', 'true'].includes(concurrent)
                && [null, 'none', 'seek'].includes(nextAction)
                && [null, 'none', 'skipTimed'].includes(previousAction)
                && (!duration || isIndefiniteTimelineTime(duration))
                && !['accel', 'decel', 'tmFilter', 'spd', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => timeNode.hasAttribute(attribute))
                && !descendants(sequenceNode, 'subTnLst').length
                && !descendants(sequenceNode, 'seq').some(nested => nested !== sequenceNode && (child(nested, 'nextCondLst') || child(nested, 'prevCondLst')))
                && advancesOnNext
        }
        const append = (
            node: Element,
            owner: Element,
            nodeStartMs: number,
            depth: number,
            fillSequenceScopes: Array<{key: string; order: number}> = [{key: sequenceKey, order: sequenceOrder}],
            timeNodeSources: Map<string, {startMs: number; endMs?: number}> = new Map(inheritedTimeNodeSources)
        ): number | undefined => {
            if (depth > 32) {
                warnings.add('A p:subTnLst timing container nesting deeper than 32 levels was skipped.')
                return undefined
            }
            if (subTimeLeafTypes.has(node.localName)) {
                const timeNode = subTimeEntryTimeNode(node)
                if (!timeNode) return undefined
                const actionStart = sequence.length
                sequence.push(...parseGroup(ownerGroup, nodeStartMs, owner, sequenceKey, sequenceOrder, [], 0, node, timeNodeSources)
                    .map(action => ({...action, fillSequenceScopes})))
                const actions = sequence.slice(actionStart)
                if (actions.some(action => action.mediaWaitForEnd && action.parentEndMs === undefined)) return undefined
                let actionEndMs = actions.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), nodeStartMs)
                if (!actions.length) {
                    const rawDuration = timeNode.getAttribute('dur')
                    const durationMs = rawDuration && !isIndefiniteTimelineTime(rawDuration) ? timelineTime(rawDuration, Number.NaN) : Number.NaN
                    const offsetStartMs = offsetStartDelay(children(child(timeNode, 'stCondLst'), 'cond'))
                    const eventStartMs = offsetStartMs === undefined
                        ? subTimeEventStartTime(timeNode, nodeStartMs, timeNodeSources)
                        : nodeStartMs + offsetStartMs
                    actionEndMs = Number.isFinite(durationMs) && eventStartMs !== undefined ? eventStartMs + durationMs : Number.NaN
                }
                const timeNodeId = timeNode.getAttribute('id')
                if (timeNodeId) {
                    const offsetStartMs = offsetStartDelay(children(child(timeNode, 'stCondLst'), 'cond'))
                    const eventStartMs = offsetStartMs === undefined
                        ? subTimeEventStartTime(timeNode, nodeStartMs, timeNodeSources)
                        : nodeStartMs + offsetStartMs
                    const actionStartMs = actions.length
                        ? Math.min(...actions.map(action => action.delayMs))
                        : eventStartMs ?? Number.POSITIVE_INFINITY
                    timeNodeSources.set(timeNodeId, {
                        startMs: actionStartMs,
                        endMs: Number.isFinite(actionEndMs) ? actionEndMs : undefined
                    })
                }
                if (actions.length) return Number.isFinite(actionEndMs) ? actionEndMs : undefined
                return Number.isFinite(actionEndMs) ? actionEndMs : undefined
            }
            if (!subTimeContainerTypes.has(node.localName)) {
                warnings.add(`An unsupported p:subTnLst child node "${node.localName}" was skipped.`)
                return undefined
            }
            const timeNode = child(node, 'cTn')
            const childTimeNodes = children(child(timeNode, 'childTnLst'))
            if (!timeNode || !childTimeNodes.length) {
                warnings.add(`An empty or malformed p:${node.localName} timing container was skipped.`)
                return undefined
            }
            let parallelTimingApplied = supportsParallelSubTimeTiming(timeNode)
            const conditions = children(child(timeNode, 'stCondLst'), 'cond')
            const conditionDelayMs = offsetStartDelay(conditions)
            const eventConditionStartMs = conditionDelayMs === undefined
                ? subTimeEventStartTime(timeNode, nodeStartMs, timeNodeSources)
                : undefined
            if (conditionDelayMs === undefined && eventConditionStartMs === undefined) {
                warnings.add(`A p:${node.localName} timing container with event-based or invalid start conditions was skipped.`)
                return undefined
            }
            if (depth > 0 && timeNode.hasAttribute('masterRel')) warnings.add('A nested p:subTnLst masterRel value is ignored; only the root entry relation is applied.')
            const sequenceControlsPresent = node.localName === 'seq' && (child(node, 'nextCondLst') || child(node, 'prevCondLst') || (node.hasAttribute('concurrent') && !['0', 'false'].includes(node.getAttribute('concurrent') || '')) || (node.hasAttribute('nextAc') && node.getAttribute('nextAc') !== 'none') || (node.hasAttribute('prevAc') && node.getAttribute('prevAc') !== 'none'))
            const conditionalNavigation = sequenceControlsPresent
                ? {
                      advancesOnNext: hasImmediateSlideCondition(node, 'nextCondLst', 'onNext'),
                      rewindsOnPrevious: hasImmediateSlideCondition(node, 'prevCondLst', 'onPrev')
                  }
                : undefined
            const conditionalSequence = depth === 0
                && node === entry
                && sequenceKey === 'main'
                && steps.indexOf(sequence) >= 0
                && sequence.length === 0
                && supportsConditionalSequence(node, timeNode, childTimeNodes.length, conditionalNavigation?.advancesOnNext === true)
            if (sequenceControlsPresent && !conditionalSequence) {
                warnings.add('Some navigation, concurrency, or seek/skip behavior on a nested p:seq is not applied.')
            }
            const containerStartMs = eventConditionStartMs ?? nodeStartMs + conditionDelayMs!
            const isSequential = node.localName === 'seq'
            const fillScopeKey = node.localName === 'seq' || node.localName === 'excl'
                ? `subtime-fill-${++subTimeFillScopeId}`
                : undefined
            let sequenceCursorMs = containerStartMs
            let containerEndMs = containerStartMs
            let unknownEnd = false
            let sequentialChildrenHaveActions = true
            let sequentialTimingApplied = false
            let flattenedParallelTimingApplied = false
            const conditionalChildActions: Array<{actions: PptxAnimation[]; startMs: number}> = []
            const containerActionStart = sequence.length
            for (const [childOrder, childNode] of childTimeNodes.entries()) {
                if (childNode.namespaceURI !== PRESENTATION_NS) {
                    warnings.add(`A non-PresentationML child of p:${node.localName} was skipped.`)
                    parallelTimingApplied = false
                    sequentialChildrenHaveActions = false
                    unknownEnd = true
                    if (isSequential) break
                    continue
                }
                const childSequenceStart = sequence.length
                const childStartMs = isSequential ? sequenceCursorMs : containerStartMs
                const childFillScopes = fillScopeKey
                    ? [...fillSequenceScopes, {key: fillScopeKey, order: childOrder}]
                    : fillSequenceScopes
                const childEndMs = append(childNode, timeNode, childStartMs, depth + 1, childFillScopes, timeNodeSources)
                const childActions = sequence.slice(childSequenceStart)
                if (conditionalSequence) {
                    conditionalChildActions.push({actions: sequence.splice(childSequenceStart), startMs: childStartMs})
                }
                if (parallelTimingApplied && (!childActions.length || childActions.some(action => !action.timingWarp))) parallelTimingApplied = false
                if (isSequential && !childActions.length) sequentialChildrenHaveActions = false
                if (childEndMs === undefined) {
                    unknownEnd = true
                    if (isSequential) {
                        warnings.add('Later p:seq children were skipped because a preceding child has no finite end time.')
                        break
                    }
                    continue
                }
                containerEndMs = Math.max(containerEndMs, childEndMs)
                if (isSequential) sequenceCursorMs = childEndMs
            }
            if (isSequential) containerEndMs = sequenceCursorMs

            const hasGroupTimingCurve = ['accel', 'decel', 'tmFilter', 'spd', 'repeatCount', 'repeatDur', 'autoRev'].some(attribute => timeNode.hasAttribute(attribute))
            const hasGroupTiming = hasGroupTimingCurve || Boolean(child(timeNode, 'endCondLst'))
            const sequentialChildActions = sequence.slice(containerActionStart)
            const parallelChildActions = sequence.slice(containerActionStart)
            if (node.localName === 'par'
                && !parallelTimingApplied
                && supportsFlattenedParallelSubTimeTiming(timeNode, parallelChildActions)
                && hasGroupTiming
                && !unknownEnd
                && Number.isFinite(containerEndMs)) {
                const groupTiming = groupTimingFor(timeNode, 'Animation group', containerEndMs - containerStartMs)
                if (groupTiming) {
                    const warpedActions = applyGroupTiming(parallelChildActions, containerStartMs, groupTiming, false, timeNode)
                    if (warpedActions.every(action => Boolean(action.timingWarp))) {
                        sequence.splice(containerActionStart, parallelChildActions.length, ...warpedActions)
                        containerEndMs = warpedActions.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), containerStartMs)
                        flattenedParallelTimingApplied = true
                    }
                }
            }
            if (isSequential && supportsSequentialSubTimeTiming(timeNode, sequentialChildActions) && hasGroupTiming && sequentialChildrenHaveActions && !unknownEnd) {
                const childDerivedDurationMs = child(timeNode, 'iterate')
                    ? undefined
                    : Number.isFinite(containerEndMs) ? containerEndMs - containerStartMs : undefined
                const groupTiming = groupTimingFor(timeNode, 'Animation group', childDerivedDurationMs)
                const childActions = sequentialChildActions
                if (groupTiming && childActions.length) {
                    const warpedActions = applyGroupTiming(childActions, containerStartMs, groupTiming, false, timeNode)
                    if (warpedActions.every(action => Boolean(action.timingWarp))) {
                        sequence.splice(containerActionStart, childActions.length, ...warpedActions)
                        containerEndMs = warpedActions.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), containerStartMs)
                        sequentialTimingApplied = true
                    }
                }
            }
            const groupTimingApplied = parallelTimingApplied || sequentialTimingApplied || flattenedParallelTimingApplied
            if (!groupTimingApplied && hasGroupTimingCurve) {
                warnings.add(`Timing curves and repeats on nested p:${node.localName} containers are rendered as one pass.`)
            }
            const rawDuration = timeNode.getAttribute('dur')
            let declaredContainerEndMs: number | undefined
            if (rawDuration && !isIndefiniteTimelineTime(rawDuration)) {
                const durationMs = timelineTime(rawDuration, Number.NaN)
                if (!Number.isFinite(durationMs)) {
                    warnings.add(`A p:${node.localName} container with an invalid duration was skipped.`)
                    return undefined
                }
                if (durationMs > 0) {
                    declaredContainerEndMs = containerStartMs + durationMs
                    if (!groupTimingApplied) {
                        containerEndMs = declaredContainerEndMs
                    }
                }
            }
            // A container's end conditions stop its active child clock; descendants inherit that stop time.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.endconditionlist?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1
            const effect = `p:${node.localName} container`
            const offsetEndConditionDurationMs = offsetEndDuration(timeNode, effect)
            const eventEndConditionDurationMs = offsetEndConditionDurationMs === undefined
                && deferredEndConditionWarnings.has(timeNode)
                ? subTimeEventEndDuration(timeNode, containerStartMs, timeNodeSources)
                : undefined
            if (eventEndConditionDurationMs !== undefined) deferredEndConditionWarnings.delete(timeNode)
            const endConditionDurationMs = offsetEndConditionDurationMs ?? eventEndConditionDurationMs
            const endConditionEndMs = endConditionDurationMs !== undefined && Number.isFinite(endConditionDurationMs)
                ? containerStartMs + endConditionDurationMs
                : undefined
            if (endConditionEndMs !== undefined) {
                containerEndMs = isIndefiniteTimelineTime(rawDuration) || (unknownEnd && declaredContainerEndMs === undefined)
                    ? endConditionEndMs
                    : Math.min(containerEndMs, endConditionEndMs)
            }
            const endSync = child(timeNode, 'endSync')
            let endSyncEndMs: number | undefined
            if (endSync) {
                // p:endSync synchronizes stopping of parallel time nodes; resolve its supported runtime-node trigger.
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.endsync?view=openxml-3.0.1
                const resolvedEndSync = subTimeConditionEventTime(timeNode, endSync, timeNodeSources)
                if (resolvedEndSync) {
                    endSyncEndMs = resolvedEndSync.eventAtMs + resolvedEndSync.delayMs
                    containerEndMs = Math.min(containerEndMs, endSyncEndMs)
                } else {
                    warnings.add(`A p:${node.localName} endSync condition could not be resolved; its natural duration was used.`)
                }
            }
            const hasParentEnd = endConditionEndMs !== undefined
                || endSyncEndMs !== undefined
                || declaredContainerEndMs !== undefined
                || !unknownEnd && (rawDuration === null || timelineTime(rawDuration, Number.NaN) === 0)
            if (hasParentEnd && Number.isFinite(containerEndMs) && !conditionalSequence) {
                for (let index = containerActionStart; index < sequence.length; index++) {
                    const action = sequence[index]!
                    sequence[index] = {
                        ...action,
                        parentEndMs: Math.min(action.parentEndMs ?? Infinity, containerEndMs),
                        ...parentFillEnd(action, timeNode, containerEndMs)
                    }
                }
            }
            const timeNodeId = timeNode.getAttribute('id')
            if (timeNodeId) {
                timeNodeSources.set(timeNodeId, {
                    startMs: containerStartMs,
                    endMs: !unknownEnd && Number.isFinite(containerEndMs) ? containerEndMs : undefined
                })
            }
            if (node.localName === 'excl') {
                if (unknownEnd || !Number.isFinite(containerEndMs) || (!groupTimingApplied && hasGroupTimingCurve)) {
                    warnings.add('A p:excl child timeline could not be paused because its end or timing is unsupported.')
                } else {
                    const exclusiveGroup = {id: `excl-${++exclusiveGroupId}`, startMs: containerStartMs, endMs: containerEndMs}
                    for (let index = containerActionStart; index < sequence.length; index++) {
                        const action = sequence[index]!
                        sequence[index] = {...action, exclusiveGroups: [...(action.exclusiveGroups || []), exclusiveGroup]}
                    }
                }
            }
            if (conditionalSequence) {
                const canSplit = !unknownEnd
                    && conditionalChildActions.length === childTimeNodes.length
                    && conditionalChildActions.every(({actions}) => actions.length > 0 && actions.every(action => action.effect !== 'media' && !action.mediaWaitForEnd && !action.mediaWaitForEndKeys?.length && !action.timingWarp && !action.exclusiveGroups?.length))
                if (canSplit) {
                    const childSteps = conditionalChildActions.map(({actions, startMs}) => actions.map(action => ({...action, delayMs: Math.max(0, action.delayMs + containerStartMs - startMs)})))
                    sequence.push(...childSteps[0]!)
                    const concurrentValue = node.getAttribute('concurrent')
                    conditionalSequencePlans.push({
                        step: sequence,
                        actions: childSteps,
                        concurrent: concurrentValue === '1' || concurrentValue === 'true',
                        nextAction: node.getAttribute('nextAc') === 'seek' ? 'seek' : 'none',
                        rewindsOnPrevious: conditionalNavigation?.rewindsOnPrevious === true
                    })
                } else {
                    sequence.push(...conditionalChildActions.flatMap(({actions}) => actions))
                    warnings.add('A nested p:seq with slide navigation conditions could not be split safely; its children use natural timing.')
                }
            }
            return unknownEnd ? undefined : containerEndMs
        }
        append(entry, parentTimeNode, startMs, 0, undefined, new Map(inheritedTimeNodeSources))
    }

    const mainTimeNode = descendants(timing, 'cTn').find(node => !isInSubTimeNodeList(node) && node.getAttribute('nodeType') === 'mainSeq')
    const mainSequence = mainTimeNode?.parentElement?.localName === 'seq' ? mainTimeNode.parentElement : undefined
    const navigationCondition = (listName: 'nextCondLst' | 'prevCondLst', eventName: 'onNext' | 'onPrev') => {
        const list = child(mainSequence, listName)
        if (!list) return undefined
        const conditions = children(list, 'cond')
        const isSupported = (condition: Element) => condition.getAttribute('evt') === eventName && (!condition.hasAttribute('delay') || (isFiniteTimelineTime(condition.getAttribute('delay')) && timelineTime(condition.getAttribute('delay'), Number.NaN) === 0)) && Boolean(firstDescendant(condition, 'sldTgt'))
        const supported = conditions.some(isSupported)
        if (conditions.some(condition => !isSupported(condition))) warnings.add(`Main animation sequence ${listName} supports only immediate slide-target ${eventName} conditions.`)
        return supported
    }
    const nextAction = mainSequence?.getAttribute('nextAc') || 'none'
    const previousAction = mainSequence?.getAttribute('prevAc') || 'none'
    const sequenceConcurrent = (sequence: Element | undefined) => {
        const value = sequence?.getAttribute('concurrent')
        if (value === null || value === undefined) return false
        if (value === '1' || value === 'true') return true
        if (value === '0' || value === 'false') return false
        warnings.add('An animation sequence has an invalid concurrent value; sequential playback is used.')
        return false
    }
    if (mainSequence && !['none', 'seek'].includes(nextAction)) warnings.add(`Main animation sequence next action "${nextAction}" is not supported.`)
    if (mainSequence && !['none', 'skipTimed'].includes(previousAction)) warnings.add(`Main animation sequence previous action "${previousAction}" is not supported.`)
    const animationSequence = mainSequence
        ? {
              concurrent: sequenceConcurrent(mainSequence),
              nextAction: (['none', 'seek'].includes(nextAction) ? nextAction : 'none') as 'none' | 'seek',
              // ECMA-376 defaults prevAc to none; skipTimed backs over sequence elements that advance by time.
              // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.sequencetimenode.previousaction?view=openxml-3.0.1
              previousAction: (['none', 'skipTimed'].includes(previousAction) ? previousAction : 'none') as 'none' | 'skipTimed',
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
                  if (isInSubTimeNodeList(node)) return false
                  const directGroup = groups.some(group => child(group, 'cTn') === node)
                  return !nestedInteractiveNode(node) && (directGroup || hasPresetClass(node))
              })
            : []
    const automaticAnimations: PptxAnimation[] = []
    const steps: PptxAnimation[][] = []
    const triggeredAnimations: NonNullable<PptxSlide['triggeredAnimations']> = []
    // p:bldLst persists text-build metadata; advAuto controls timed build advancement and grpId links it to p:cTn.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.buildparagraph?view=openxml-3.0.1
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/2936b07f-699e-42a7-a898-70122183e0a2
    const paragraphBuilds = children(child(timing, 'bldLst'), 'bldP')
    const timeNodeSources = new Map<string, {step?: number; startDelayMs?: number; endDelayMs?: number; mediaActionKey?: string}>()
    const eventTimeNodeSourcesForStep = <T extends {step?: number; startDelayMs?: number; endDelayMs?: number}>(
        sources: Map<string, T>,
        step: number | undefined
    ) => {
        const result = new Map<string, {startMs: number; endMs?: number}>()
        for (const [id, source] of sources) {
            if (source.step !== step || source.startDelayMs === undefined) continue
            result.set(id, {
                startMs: source.startDelayMs,
                ...(source.endDelayMs === undefined ? {} : {endMs: source.endDelayMs})
            })
        }
        return result
    }
    const mainSequenceId = mainTimeNode?.getAttribute('id')
    if (mainSequenceId) timeNodeSources.set(mainSequenceId, {startDelayMs: 0})
    const buildAutoAdvanceForGroup = (group: Element, timeNode: Element, actions: PptxAnimation[]) => {
        if (!paragraphBuilds.length || !actions.length) return undefined
        const groupIds = new Set<string>()
        let owner: Element | null = timeNode
        while (owner && (owner === group || group.contains(owner))) {
            if (owner.localName === 'cTn') {
                const groupId = owner.getAttribute('grpId')
                if (groupId) groupIds.add(groupId)
            }
            if (owner === group) break
            owner = owner.parentElement
        }
        const targetIds = new Set(actions.map(action => action.targetId))
        const matches = paragraphBuilds.filter(build => {
            const shapeId = build.getAttribute('spid')
            const groupId = build.getAttribute('grpId')
            if (!shapeId || !targetIds.has(shapeId) || !groupId) return false
            return groupIds.has(groupId)
        })
        if (!matches.length) return undefined
        const rawValues = matches.map(build => trimXmlWhitespace(build.getAttribute('advAuto')) ?? null)
        if (new Set(rawValues).size > 1) {
            warnings.add('Paragraph builds in one animation step specify different auto-advance times; automatic build progression was ignored.')
            return undefined
        }
        const rawValue = rawValues[0]
        if (rawValue === null || rawValue === undefined) return undefined
        if (rawValue === 'indefinite') return 0
        if (!/^\d+$/.test(rawValue)) {
            warnings.add('A paragraph build auto-advance time was invalid and ignored.')
            return undefined
        }
        const delayMs = Number(rawValue)
        if (!Number.isSafeInteger(delayMs) || delayMs > MAX_TL_TIME_MS) {
            warnings.add('A paragraph build auto-advance time exceeds the supported PowerPoint range and was ignored.')
            return undefined
        }
        return delayMs
    }
    let currentStep: PptxAnimation[] | undefined
    for (const [sequenceOrder, timeNode] of mainEvents.entries()) {
        const group = mainGroupFor(timeNode)
        if (!group) continue
        const type = groupType(timeNode)
        if (startsOnClick(type)) {
            currentStep = []
            steps.push(currentStep)
        } else if (!['withEffect', 'afterEffect', 'withGroup', 'afterGroup'].includes(type || '')) {
            continue
        }
        const sequence = currentStep || automaticAnimations
        const mediaWaitForEndKeys = [...new Set(sequence.flatMap(action => (action.mediaWaitForEnd && action.mediaActionKey ? [action.mediaActionKey] : [])))]
        const precedingEnd = followsPreviousGroup(type) ? sequence.reduce((end, action) => {
            const duration = playbackDuration(action)
            // An indefinite media node completes from its media event, represented by mediaWaitForEndKeys.
            // Keep it out of the numeric offset so an afterEffect can be released when that event arrives.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonmedianode?view=openxml-3.0.1
            if (action.effect === 'media' && action.mediaWaitForEnd && action.mediaActionKey && !Number.isFinite(duration)) return end
            return Math.max(end, action.delayMs + duration)
        }, 0) : 0
        const eventTimeNodeSources = eventTimeNodeSourcesForStep(timeNodeSources, currentStep ? steps.length - 1 : undefined)
        const parentTimeNodes: Element[] = []
        let parent: Element | null = timeNode
        while (parent && parent !== mainTimeNode) {
            if (parent.localName === 'cTn') parentTimeNodes.push(parent)
            parent = parent.parentElement
        }
        let resolvedGroupStartMs = precedingEnd
        for (const parentTimeNode of parentTimeNodes.reverse()) {
            // Main-sequence clickEffect/clickPar nodes use a lone indefinite delay as the
            // click-build activation sentinel. The player activates that build on user input,
            // so nested withEffect nodes share the build boundary instead of inheriting Infinity.
            // Keep unrelated indefinite conditions unresolved.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timenodevalues?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.startconditionlist?view=openxml-3.0.1
            const startMs = isClickBuildWait(parentTimeNode, groupType(parentTimeNode))
                ? resolvedGroupStartMs
                : subTimeEventStartTime(parentTimeNode, resolvedGroupStartMs, eventTimeNodeSources)
            if (startMs === undefined) {
                warnings.add('An animation with event-based or invalid start conditions could not be activated.')
                resolvedGroupStartMs = Number.POSITIVE_INFINITY
                break
            }
            resolvedGroupStartMs = startMs
        }
        const groupStartDelay = resolvedGroupStartMs - precedingEnd
        const startDelayMs = resolvedGroupStartMs
        const actions = parseGroup(group, startDelayMs, timeNode, 'main', sequenceOrder, followsPreviousGroup(type) ? mediaWaitForEndKeys : [], groupStartDelay, undefined, eventTimeNodeSources)
        const buildAutoAdvanceAfterMs = currentStep ? buildAutoAdvanceForGroup(group, timeNode, actions) : undefined
        sequence.push(...(buildAutoAdvanceAfterMs === undefined
            ? actions
            : actions.map(action => ({...action, buildAutoAdvanceAfterMs}))))
        const timeNodeId = timeNode.getAttribute('id')
        if (timeNodeId) {
            const waitsForMedia = actions.some(action => action.mediaWaitForEndKeys?.length)
            const hasMedia = actions.some(action => action.effect === 'media')
            const rootMediaActionKey = actions.find(action => action.sourceTimeNodeId === timeNodeId && action.effect === 'media')?.mediaActionKey
            const endDelayMs = actions.length && !waitsForMedia && !hasMedia
                ? parsedTimeNodeEndOffsets.get(timeNode) ?? Math.max(startDelayMs, ...actions.map(action => action.delayMs + playbackDuration(action)))
                : undefined
            timeNodeSources.set(timeNodeId, {
                step: currentStep ? steps.length - 1 : undefined,
                startDelayMs: waitsForMedia ? undefined : startDelayMs,
                endDelayMs,
                ...(rootMediaActionKey ? {mediaActionKey: rootMediaActionKey} : {})
            })
        }
        const actionsByTimeNodeId = new Map<string, PptxAnimation[]>()
        for (const action of actions) {
            const sourceTimeNodeId = action.sourceTimeNodeId
            if (!sourceTimeNodeId || sourceTimeNodeId === timeNodeId || action.effect === 'media') continue
            const sourceActions = actionsByTimeNodeId.get(sourceTimeNodeId) || []
            sourceActions.push(action)
            actionsByTimeNodeId.set(sourceTimeNodeId, sourceActions)
        }
        for (const [sourceTimeNodeId, sourceActions] of actionsByTimeNodeId) {
            if (timeNodeSources.has(sourceTimeNodeId)) continue
            const endDelays = sourceActions.map(action => action.delayMs + playbackDuration(action))
            timeNodeSources.set(sourceTimeNodeId, {
                step: currentStep ? steps.length - 1 : undefined,
                startDelayMs: Math.min(...sourceActions.map(action => action.delayMs)),
                endDelayMs: endDelays.every(Number.isFinite) ? Math.max(...endDelays) : undefined
            })
        }
    }
    for (const list of descendants(timing, 'subTnLst')) {
        const owner = list.parentElement
        const ownerId = owner?.namespaceURI === PRESENTATION_NS && owner.localName === 'cTn' ? owner.getAttribute('id') : null
        const source = ownerId ? timeNodeSources.get(ownerId) : undefined
        const ownerGroup = owner && mainEvents.includes(owner) ? mainGroupFor(owner) : undefined
        if (list.namespaceURI !== PRESENTATION_NS || !owner || !source || !ownerGroup) {
            continue
        }
        for (const entry of children(list)) {
            if (entry.namespaceURI !== PRESENTATION_NS || !subTimeEntryTypes.has(entry.localName)) continue
            const timeNode = subTimeEntryTimeNode(entry)
            const masterRelation = timeNode?.getAttribute('masterRel')
            const afterEffectValue = timeNode?.getAttribute('afterEffect')
            // ECMA-376 defines lastClick, but PowerPoint does not support it.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/3de95cb4-776a-4afc-a195-66a1470dc452
            if (masterRelation === 'lastClick') {
                warnings.add('PowerPoint does not support p:cTn@masterRel="lastClick"; the sub-time entry was skipped.')
                continue
            }
            if (!timeNode || !['sameClick', 'nextClick'].includes(masterRelation || '') || (afterEffectValue !== null && afterEffectValue !== undefined && !['0', '1', 'false', 'true'].includes(afterEffectValue))) {
                continue
            }
            const targetStep = masterRelation === 'sameClick'
                ? source.step
                : (source.step ?? -1) + 1
            const matchingClickEvents = targetStep === undefined ? [] : mainEvents.filter(node =>
                startsOnClick(groupType(node)) && node.getAttribute('id') && timeNodeSources.get(node.getAttribute('id')!)?.step === targetStep)
            const targetEvent = masterRelation === 'sameClick'
                ? owner
                : matchingClickEvents[0]
            if (!targetEvent) continue
            const targetSourceId = targetEvent.getAttribute('id')
            const targetSource = targetSourceId ? timeNodeSources.get(targetSourceId) : undefined
            const sequence = targetStep === undefined ? automaticAnimations : steps[targetStep]
            const delayMs = afterEffectValue === '1' || afterEffectValue === 'true' ? targetSource?.endDelayMs : targetSource?.startDelayMs
            if (!targetSource || !sequence || delayMs === undefined || !Number.isFinite(delayMs)) continue
            const targetOrder = mainEvents.indexOf(targetEvent)
            appendSubTimeEntry(entry, ownerGroup, owner, sequence, 'main', targetOrder + 0.5, delayMs, eventTimeNodeSourcesForStep(timeNodeSources, targetStep))
            handledSubTimeEntries.add(entry)
        }
    }
    const interactiveSequences = descendants(timing, 'cTn').filter(node => !isInSubTimeNodeList(node) && node.getAttribute('nodeType') === 'interactiveSeq')
    const skippedInteractiveSequences = new Set<Element>()
    const isInSkippedInteractiveSequence = (node: Element) => {
        let parent: Element | null = node
        while (parent && parent !== timing) {
            if (parent.localName === 'cTn' && parent.getAttribute('nodeType') === 'interactiveSeq' && skippedInteractiveSequences.has(parent)) return true
            parent = parent.parentElement
        }
        return false
    }
    const shapeTriggerEvents = new Set(['onClick', 'onDblClick', 'onMouseOver', 'onMouseOut'])
    // PowerPoint defaults cTn@restart to `always`; restart state is modeled for interactive sequences.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/02b2312e-9700-42e0-8a19-b2af1ba83a20
    for (const timeNode of descendants(timing, 'cTn')) {
        if (isInSubTimeNodeList(timeNode)) continue
        const restart = timeNode.getAttribute('restart')
        if (!restart || restart === 'always') continue
        // restart only controls whether a time node may restart after it completes its action.
        // A tmRoot with an indefinite duration cannot complete, so its restart policy has no playback effect.
        // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/02b2312e-9700-42e0-8a19-b2af1ba83a20
        const isIndefiniteTimelineRoot = timeNode.getAttribute('nodeType') === 'tmRoot' && isIndefiniteTimelineTime(timeNode.getAttribute('dur'))
        if (isIndefiniteTimelineRoot && ['whenNotActive', 'never'].includes(restart)) continue
        if (timeNode.getAttribute('nodeType') === 'interactiveSeq' && ['whenNotActive', 'never'].includes(restart)) continue
        warnings.add(['whenNotActive', 'never'].includes(restart)
            ? 'A non-default p:cTn@restart outside an interactive sequence is not applied yet.'
            : 'An invalid p:cTn@restart uses the PowerPoint default "always".')
    }
    // PowerPoint distinguishes begin/end time-node events from onBegin/onEnd events on p:cond's target.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/6acf6d2c-6423-4f5e-8a16-6533a20fbd41
    // Resolve non-media onBegin/onEnd only for parsed action nodes with runtime event times; media nodes use playback events.
    for (const [interactiveIndex, interactive] of interactiveSequences.entries()) {
        const conditions = children(child(interactive, 'stCondLst'), 'cond')
        const sequenceContainer = interactive.parentElement?.localName === 'seq' ? interactive.parentElement : undefined
        const restart = interactive.getAttribute('restart')
        const conditionWarnings = new Set<string>()
        const triggerForCondition = (condition: Element): PptxAnimationTriggerCondition | undefined => {
            const event = condition.getAttribute('evt') || ''
            const rawDelay = condition.getAttribute('delay')
            const delay = trimXmlWhitespace(rawDelay)
            if (isIndefiniteTimelineTime(delay)) {
                conditionWarnings.add('An interactive animation with an indefinite start-condition delay was skipped.')
                return undefined
            }
            const delayMs = rawDelay === null ? 0 : delay && isFiniteTimelineTime(delay) ? timelineTime(delay, Number.NaN) : Number.NaN
            if (!Number.isFinite(delayMs)) {
                conditionWarnings.add('An interactive animation with an invalid start-condition delay was skipped.')
                return undefined
            }
            const shapeTarget = firstDescendant(condition, 'spTgt')
            const bookmarkTarget = firstDescendant(condition, 'bmkTgt')
            if (event === 'onMediaBookmark' && bookmarkTarget) {
                const id = bookmarkTarget.getAttribute('spid') || ''
                const bookmarkName = bookmarkTarget.getAttribute('bmkName') || ''
                const mediaCandidate = elementsById.get(id)
                const media = mediaCandidate && (mediaCandidate.kind === 'audio' || mediaCandidate.kind === 'video') && mediaCandidate.mediaUrl ? mediaCandidate : undefined
                const bookmark = media?.mediaBookmarks?.find(item => item.name === bookmarkName)
                return bookmark && id && bookmarkName ? {type: 'mediaBookmark', id, bookmarkName, timeMs: bookmark.timeMs, delayMs} : undefined
            }
            if (shapeTriggerEvents.has(event) && shapeTarget) {
                const id = shapeTarget.getAttribute('spid') || ''
                return id && renderedElementIds.has(id) ? {type: 'shape', id, event: event as 'onClick' | 'onDblClick' | 'onMouseOver' | 'onMouseOut', delayMs} : undefined
            }
            if ((event === 'onBegin' || event === 'onEnd') && shapeTarget) {
                const diagramId = shapeTarget.getAttribute('spid') || ''
                const subShapeId = child(shapeTarget, 'subSp')?.getAttribute('spid')
                const id = subShapeId ? smartArtAnimationTargets.get(`${diagramId}:${subShapeId}`) || '' : diagramId
                return id && renderedElementIds.has(id) ? {type: 'animatedTarget', id, event, delayMs} : undefined
            }
            if ((event === 'onNext' || event === 'onPrev' || event === 'onStopAudio') && firstDescendant(condition, 'sldTgt')) {
                return {type: 'slide', event, delayMs}
            }
            const timeNodeId = firstDescendant(condition, 'tn')?.getAttribute('val')
            const isBeginEvent = event === 'begin' || event === 'onBegin'
            const isEndEvent = event === 'end' || event === 'onEnd'
            if ((isBeginEvent || isEndEvent) && timeNodeId) {
                const source = timeNodeSources.get(timeNodeId)
                const sourceDelayMs = isBeginEvent ? source?.startDelayMs : source?.endDelayMs
                const mediaTimeNodeEvent = source?.mediaActionKey !== undefined
                if (source && ((sourceDelayMs !== undefined && Number.isFinite(sourceDelayMs)) || mediaTimeNodeEvent)) {
                    return {
                        type: 'timeNode',
                        id: timeNodeId,
                        event: event as 'begin' | 'onBegin' | 'end' | 'onEnd',
                        sourceStep: source.step,
                        sourceDelayMs: sourceDelayMs ?? 0,
                        ...(mediaTimeNodeEvent ? {sourceMediaActionKey: source.mediaActionKey} : {}),
                        delayMs
                    }
                }
            }
            return undefined
        }
        const triggers = conditions.flatMap(condition => {
            const trigger = triggerForCondition(condition)
            return trigger ? [trigger] : []
        })
        if (!conditions.length || triggers.length !== conditions.length) {
            skippedInteractiveSequences.add(interactive)
            if (conditionWarnings.size) conditionWarnings.forEach(warning => warnings.add(warning))
            else warnings.add(conditions.length
                ? 'An interactive animation with an unsupported start condition was skipped.'
                : 'An interactive animation without a supported shape, animated-target, slide navigation, timing-node, media-bookmark, or onStopAudio trigger was skipped.')
            continue
        }

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
        const interactiveTimeNodeSources = new Map<string, {step: number; startDelayMs?: number; endDelayMs?: number; mediaActionKey?: string}>()
        const markerOrders = new Map<Element, number>()
        let currentStep: PptxAnimation[] | undefined
        for (const [sequenceOrder, timeNode] of markers.entries()) {
            const type = groupType(timeNode)
            if (startsOnClick(type) || !currentStep) {
                currentStep = []
                triggerSteps.push(currentStep)
            }
            const group = timeNode.parentElement
            if (!group) continue
            const mediaWaitForEndKeys = [...new Set(currentStep.flatMap(action => (action.mediaWaitForEnd && action.mediaActionKey ? [action.mediaActionKey] : [])))]
            const precedingEnd = followsPreviousGroup(type) ? currentStep.reduce((end, action) => Math.max(end, action.delayMs + playbackDuration(action)), 0) : 0
            const eventTimeNodeSources = eventTimeNodeSourcesForStep(interactiveTimeNodeSources, triggerSteps.length - 1)
            const parentTimeNodes: Element[] = []
            let parent: Element | null = timeNode
            while (parent && parent !== interactive) {
                if (parent.localName === 'cTn') parentTimeNodes.push(parent)
                parent = parent.parentElement
            }
            let resolvedGroupStartMs = precedingEnd
            for (const parentTimeNode of parentTimeNodes.reverse()) {
                // Interactive clickEffect/clickPar nodes are already dispatched by the shape
                // trigger above; map a parent click node's lone indefinite activation sentinel
                // to this step's local clock, including when parsing a nested withEffect child.
                const startMs = isClickBuildWait(parentTimeNode, groupType(parentTimeNode))
                    ? resolvedGroupStartMs
                    : subTimeEventStartTime(parentTimeNode, resolvedGroupStartMs, eventTimeNodeSources)
                if (startMs === undefined) {
                    warnings.add('An interactive animation with event-based or invalid start conditions could not be activated.')
                    resolvedGroupStartMs = Number.POSITIVE_INFINITY
                    break
                }
                resolvedGroupStartMs = startMs
            }
            const groupStartDelay = resolvedGroupStartMs - precedingEnd
            const startDelayMs = resolvedGroupStartMs
            const actions = parseGroup(group, startDelayMs, timeNode, `interactive:${interactiveIndex}`, sequenceOrder, followsPreviousGroup(type) ? mediaWaitForEndKeys : [], groupStartDelay, undefined, eventTimeNodeSources)
            const buildAutoAdvanceAfterMs = buildAutoAdvanceForGroup(group, timeNode, actions)
            currentStep.push(...(buildAutoAdvanceAfterMs === undefined
                ? actions
                : actions.map(action => ({...action, buildAutoAdvanceAfterMs}))))
            markerOrders.set(timeNode, sequenceOrder)
            const timeNodeId = timeNode.getAttribute('id')
            if (timeNodeId) {
                const waitsForMedia = actions.some(action => action.mediaWaitForEndKeys?.length)
                const hasMedia = actions.some(action => action.effect === 'media')
                const rootMediaActionKey = actions.find(action => action.sourceTimeNodeId === timeNodeId && action.effect === 'media')?.mediaActionKey
                const endDelayMs = actions.length && !waitsForMedia && !hasMedia
                    ? parsedTimeNodeEndOffsets.get(timeNode) ?? Math.max(startDelayMs, ...actions.map(action => action.delayMs + playbackDuration(action)))
                    : undefined
                interactiveTimeNodeSources.set(timeNodeId, {
                    step: triggerSteps.length - 1,
                    startDelayMs: waitsForMedia ? undefined : startDelayMs,
                    endDelayMs,
                    ...(rootMediaActionKey ? {mediaActionKey: rootMediaActionKey} : {})
                })
            }
            const localTimeNodesById = new Map(descendants(group, 'cTn').flatMap(node => {
                const id = node.getAttribute('id')
                return id ? [[id, node] as const] : []
            }))
            const actionsByTimeNodeId = new Map<string, PptxAnimation[]>()
            for (const action of actions) {
                const sourceTimeNodeId = action.sourceTimeNodeId
                if (!sourceTimeNodeId || sourceTimeNodeId === timeNodeId) continue
                const sourceActions = actionsByTimeNodeId.get(sourceTimeNodeId) || []
                sourceActions.push(action)
                actionsByTimeNodeId.set(sourceTimeNodeId, sourceActions)
            }
            for (const [sourceTimeNodeId, sourceActions] of actionsByTimeNodeId) {
                const mediaActionKey = sourceActions.find(action => action.effect === 'media')?.mediaActionKey
                if (mediaActionKey) {
                    interactiveTimeNodeSources.set(sourceTimeNodeId, {step: triggerSteps.length - 1, mediaActionKey})
                    continue
                }
                if (interactiveTimeNodeSources.has(sourceTimeNodeId)) continue
                const timedActions = sourceActions.filter(action => action.effect !== 'media')
                if (!timedActions.length) continue
                const sourceTimeNode = localTimeNodesById.get(sourceTimeNodeId)
                const startDelayMs = Math.min(...timedActions.map(action => action.delayMs))
                const endDelayMs = sourceTimeNode ? parsedTimeNodeEndOffsets.get(sourceTimeNode) : undefined
                const naturalEndDelayMs = Math.max(startDelayMs, ...timedActions.map(action => action.delayMs + playbackDuration(action)))
                interactiveTimeNodeSources.set(sourceTimeNodeId, {
                    step: triggerSteps.length - 1,
                    startDelayMs,
                    endDelayMs: endDelayMs ?? naturalEndDelayMs
                })
            }
        }
        for (const list of descendants(interactive, 'subTnLst')) {
            const owner = list.parentElement
            const ownerId = owner?.namespaceURI === PRESENTATION_NS && owner.localName === 'cTn' ? owner.getAttribute('id') : null
            const source = ownerId ? interactiveTimeNodeSources.get(ownerId) : undefined
            const ownerOrder = owner ? markerOrders.get(owner) : undefined
            const group = owner?.parentElement
            if (list.namespaceURI !== PRESENTATION_NS || !owner || !source || ownerOrder === undefined || !group) continue
            for (const entry of children(list)) {
                if (entry.namespaceURI !== PRESENTATION_NS || !subTimeEntryTypes.has(entry.localName)) continue
                const timeNode = subTimeEntryTimeNode(entry)
                const masterRelation = timeNode?.getAttribute('masterRel')
                const afterEffectValue = timeNode?.getAttribute('afterEffect')
                if (masterRelation === 'lastClick') {
                    warnings.add('PowerPoint does not support p:cTn@masterRel="lastClick"; the sub-time entry was skipped.')
                    continue
                }
                if (!timeNode || !['sameClick', 'nextClick'].includes(masterRelation || '') || (afterEffectValue !== null && afterEffectValue !== undefined && !['0', '1', 'false', 'true'].includes(afterEffectValue))) continue
                const targetStep = masterRelation === 'sameClick'
                    ? source.step
                    : source.step + 1
                const matchingClickMarkers = markers.filter(node => {
                    const nodeId = node.getAttribute('id')
                    return startsOnClick(groupType(node)) && Boolean(nodeId) && interactiveTimeNodeSources.get(nodeId!)?.step === targetStep
                })
                const targetEvent = masterRelation === 'sameClick'
                    ? owner
                    : matchingClickMarkers[0]
                if (!targetEvent) continue
                const targetId = targetEvent.getAttribute('id')
                const targetSource = targetId ? interactiveTimeNodeSources.get(targetId) : undefined
                const sequence = triggerSteps[targetStep]
                const delayMs = afterEffectValue === '1' || afterEffectValue === 'true' ? targetSource?.endDelayMs : targetSource?.startDelayMs
                const targetOrder = markerOrders.get(targetEvent)
                if (!targetSource || !sequence || delayMs === undefined || !Number.isFinite(delayMs) || targetOrder === undefined) continue
                appendSubTimeEntry(entry, group, owner, sequence, `interactive:${interactiveIndex}`, targetOrder + 0.5, delayMs, eventTimeNodeSourcesForStep(interactiveTimeNodeSources, targetStep))
                handledSubTimeEntries.add(entry)
            }
        }
        const interactiveEndSync = child(interactive, 'endSync')
        if (interactiveEndSync) {
            const runtimeTrigger = interactiveEndSync.children.length === 1
                && interactiveEndSync.children[0]?.namespaceURI === PRESENTATION_NS
                ? interactiveEndSync.children[0]
                : undefined
            const event = interactiveEndSync.getAttribute('evt')
            const rawDelay = interactiveEndSync.getAttribute('delay')
            const delayMs = rawDelay === null ? 0 : isFiniteTimelineTime(rawDelay) ? timelineTime(rawDelay, Number.NaN) : Number.NaN
            let resolvedTrigger: {step: number; eventAtMs: number} | undefined
            const validEndSyncAttributes = Array.from(interactiveEndSync.attributes).every(attribute => ['evt', 'delay'].includes(attribute.localName))
            if (validEndSyncAttributes && runtimeTrigger && (event === 'begin' || event === 'end') && Number.isFinite(delayMs)) {
                const eventTimeForSource = (source: {step: number; startDelayMs?: number; endDelayMs?: number}) => {
                    const eventAtMs = event === 'begin' ? source.startDelayMs : source.endDelayMs
                    return eventAtMs !== undefined && Number.isFinite(eventAtMs) ? {step: source.step, eventAtMs} : undefined
                }
                if (runtimeTrigger.localName === 'tn') {
                    const sourceId = runtimeTrigger.getAttribute('val')
                    const source = sourceId ? interactiveTimeNodeSources.get(sourceId) : undefined
                    resolvedTrigger = source ? eventTimeForSource(source) : undefined
                } else if (runtimeTrigger.localName === 'rtn') {
                    const triggerValue = runtimeTrigger.getAttribute('val')
                    const childTimeNodes = children(child(interactive, 'childTnLst'))
                        .filter(node => node.namespaceURI === PRESENTATION_NS)
                        .map(node => child(node, 'cTn'))
                    const childSources = childTimeNodes.map(timeNode => {
                        const id = timeNode?.getAttribute('id')
                        return id ? interactiveTimeNodeSources.get(id) : undefined
                    })
                    const allChildSourcesResolved = childSources.length > 0
                        && childSources.every(source => source !== undefined)
                        && new Set(childSources.map(source => source?.step)).size === 1
                    if (allChildSourcesResolved && ['first', 'last', 'all'].includes(triggerValue || '')) {
                        const resolvedSources = childSources as Array<{step: number; startDelayMs?: number; endDelayMs?: number}>
                        const resolvedChildEvents = resolvedSources.map(eventTimeForSource)
                        if (resolvedChildEvents.every(source => source !== undefined)) {
                            const events = resolvedChildEvents as Array<{step: number; eventAtMs: number}>
                            if (triggerValue === 'all') {
                                resolvedTrigger = {step: events[0]!.step, eventAtMs: Math.max(...events.map(source => source.eventAtMs))}
                            } else if (resolvedSources.every(source => Number.isFinite(source.endDelayMs))) {
                                const orderedSources = [...resolvedSources].sort((left, right) => (left.endDelayMs ?? Infinity) - (right.endDelayMs ?? Infinity))
                                const selectedSource = triggerValue === 'first' ? orderedSources[0] : orderedSources[orderedSources.length - 1]
                                const selected = selectedSource ? eventTimeForSource(selectedSource) : undefined
                                if (selected) resolvedTrigger = selected
                            }
                        }
                    }
                }
            }
            if (resolvedTrigger) {
                const endAtMs = resolvedTrigger.eventAtMs + delayMs
                const stepActions = triggerSteps[resolvedTrigger.step]
                if (stepActions && Number.isFinite(endAtMs)) {
                    const interactiveTimeNode = child(interactive, 'cTn')
                    triggerSteps[resolvedTrigger.step] = stepActions.map(action => ({
                        ...action,
                        parentEndMs: Math.min(action.parentEndMs ?? Infinity, endAtMs),
                        ...parentFillEnd(action, interactiveTimeNode, endAtMs)
                    }))
                } else {
                    warnings.add('An interactive p:endSync condition could not be applied; its natural duration was used.')
                }
            } else {
                warnings.add('An interactive p:endSync condition could not be resolved; its natural duration was used.')
            }
        }
        if (triggerSteps.some(step => step.length)) {
            triggeredAnimations.push({
                conditions: triggers,
                steps: triggerSteps,
                concurrent: sequenceConcurrent(sequenceContainer),
                ...(restart === 'whenNotActive' || restart === 'never' ? {restart} : {})
            })
        }
    }
    if (descendants(timing, 'subTnLst').some(list => list.namespaceURI === PRESENTATION_NS && children(list).some(entry => !handledSubTimeEntries.has(entry)))) {
        warnings.add('Animations in non-empty p:subTnLst lists use unsupported master-relative timing and may be missing dim-after or sound-effect behaviors.')
    }
    let animationStepSequences: PptxSlide['animationStepSequences']
    if (conditionalSequencePlans.length) {
        const plansByStep = new Map<number, typeof conditionalSequencePlans>()
        for (const plan of conditionalSequencePlans) {
            const stepIndex = steps.indexOf(plan.step)
            if (stepIndex < 0) continue
            const plans = plansByStep.get(stepIndex) || []
            plans.push(plan)
            plansByStep.set(stepIndex, plans)
        }
        const expandedSteps: PptxAnimation[][] = []
        const expandedStepSequences: NonNullable<PptxSlide['animationStepSequences']> = []
        const remappedSteps = new Map<number, number>()
        for (let stepIndex = 0; stepIndex < steps.length; stepIndex++) {
            const baseExpandedStepIndex = expandedSteps.length
            remappedSteps.set(stepIndex, baseExpandedStepIndex)
            expandedSteps.push(steps[stepIndex]!)
            const plans = plansByStep.get(stepIndex) || []
            const childCount = Math.max(0, ...plans.map(plan => plan.actions.length))
            const sequenceMetadata = (childIndex: number) => {
                const advancesOnNext = plans.some(plan => plan.actions.length > childIndex + 1)
                const hasPreviousBoundary = childIndex > 0 && plans.some(plan => plan.actions.length > childIndex)
                const rewindsWithinSequence = hasPreviousBoundary && plans.some(plan => plan.rewindsOnPrevious && plan.actions.length > childIndex)
                const rewindsToParentSequence = hasPreviousBoundary && !rewindsWithinSequence && animationSequence?.rewindsOnPrevious !== false
                const rewindsOnPrevious = rewindsWithinSequence || rewindsToParentSequence
                const advancingPlan = advancesOnNext ? plans.find(plan => plan.actions.length > childIndex + 1) : undefined
                return advancesOnNext || hasPreviousBoundary
                    ? {
                          ...(advancesOnNext ? {concurrent: advancingPlan!.concurrent, nextAction: advancingPlan!.nextAction, advancesOnNext: true} : {}),
                          ...(hasPreviousBoundary
                              ? {
                                    rewindsOnPrevious,
                                    ...(rewindsOnPrevious
                                        ? {previousStepIndex: rewindsWithinSequence ? baseExpandedStepIndex + childIndex - 1 : baseExpandedStepIndex - 1}
                                        : {})
                                }
                              : {})
                      }
                    : undefined
            }
            expandedStepSequences.push(childCount ? sequenceMetadata(0) : undefined)
            for (let childIndex = 1; childIndex < childCount; childIndex++) {
                expandedSteps.push(plans.flatMap(plan => plan.actions[childIndex] || []))
                expandedStepSequences.push(sequenceMetadata(childIndex))
            }
        }
        steps.splice(0, steps.length, ...expandedSteps)
        animationStepSequences = expandedStepSequences.some(Boolean) ? expandedStepSequences : undefined
        for (const sequence of triggeredAnimations) {
            for (const condition of sequence.conditions) {
                if (condition.type !== 'timeNode' || condition.sourceStep === undefined) continue
                condition.sourceStep = remappedSteps.get(condition.sourceStep) ?? condition.sourceStep
            }
        }
    }
    if (steps.some(step => new Set(step.flatMap(action => action.buildAutoAdvanceAfterMs === undefined ? [] : [action.buildAutoAdvanceAfterMs])).size > 1)) {
        warnings.add('A click step contains paragraph builds with different auto-advance times; automatic build progression was ignored.')
    }
    if (triggeredAnimations.some(sequence => sequence.steps.some(step => new Set(step.flatMap(action => action.buildAutoAdvanceAfterMs === undefined ? [] : [action.buildAutoAdvanceAfterMs])).size > 1))) {
        warnings.add('An interactive animation step contains paragraph builds with different auto-advance times; automatic build progression was ignored.')
    }
    const timelineContexts = [automaticAnimations, ...steps, ...triggeredAnimations.flatMap(sequence => sequence.steps)]
    const allTimelineActions = timelineContexts.flat()
    const exclusiveGroups = new Map<string, {startMs: number; endMs: number}>()
    for (const action of allTimelineActions) {
        for (const group of action.exclusiveGroups || []) exclusiveGroups.set(group.id, group)
    }
    let mediaMayContinueAfterExclusive = false
    for (const [groupId, group] of exclusiveGroups) {
        const owner = allTimelineActions.find(action => action.exclusiveGroups?.some(candidate => candidate.id === groupId))
        const ownerContext = owner ? timelineContexts.find(context => context.includes(owner)) : undefined
        if (owner && allTimelineActions.some(action => action.effect === 'media'
            && action.exclusiveGroups?.some(candidate => candidate.id === groupId)
            && ownerContext?.includes(action)
            && action.delayMs + playbackDuration(action) > group.endMs)) {
            mediaMayContinueAfterExclusive = true
            break
        }
    }
    if (mediaMayContinueAfterExclusive) {
        warnings.add('Media started inside p:excl may continue after the exclusive timeline ends.')
    }
    if (descendants(timing, 'animEffect').some(effect => !isInSubTimeNodeList(effect) && !handledEffects.has(effect) && !isInSkippedInteractiveSequence(effect))) {
        warnings.add('An object effect animation outside supported timing groups is not rendered yet.')
    }
    if (descendants(timing, 'animScale').some(animation => !isInSubTimeNodeList(animation) && !handledScales.has(animation) && !isInSkippedInteractiveSequence(animation))) {
        warnings.add('Scale animations outside supported click/automatic groups are not rendered yet.')
    }
    if (descendants(timing, 'animMotion').some(animation => !isInSubTimeNodeList(animation) && !handledMotions.has(animation) && !isInSkippedInteractiveSequence(animation))) {
        warnings.add('Object motion, emphasis, and property animations are not rendered yet.')
    }
    if (descendants(timing, 'animRot').some(animation => !isInSubTimeNodeList(animation) && !handledRotations.has(animation) && !isInSkippedInteractiveSequence(animation))) {
        warnings.add('Unsupported rotation animations are not rendered yet.')
    }
    if (descendants(timing, 'animClr').some(animation => !isInSubTimeNodeList(animation) && !handledColors.has(animation) && !isInSkippedInteractiveSequence(animation))) {
        warnings.add('Unsupported color animations are not rendered yet.')
    }
    if (descendants(timing, 'anim').some(animation => !isInSubTimeNodeList(animation) && !handledTextAnimations.has(animation) && !isInSkippedInteractiveSequence(animation))) {
        warnings.add('Unsupported generic motion, emphasis, and property animations are not rendered yet.')
    }
    if (['audio', 'video', 'cmd'].some(name => descendants(timing, name).some(node => node.namespaceURI === PRESENTATION_NS && !isInSubTimeNodeList(node) && !handledMediaActions.has(node) && !isInSkippedInteractiveSequence(node)))) {
        warnings.add('Media playback outside supported slide timing groups is not rendered yet.')
    }
    for (const effect of deferredEndConditionWarnings.values()) {
        warnings.add(`${effect} has an event-based end condition that could not be resolved; its declared timing was used.`)
    }
    return {
        automaticAnimations: automaticAnimations.length ? automaticAnimations : undefined,
        animationSteps: steps.length ? steps : undefined,
        animationSequence,
        animationStepSequences,
        triggeredAnimations: triggeredAnimations.length ? triggeredAnimations : undefined
    }
}

async function parsePicture(shape: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>, fillStyles?: ParsedTheme): Promise<PptxElement> {
    const properties = child(shape, 'spPr')
    const pictureEffects = referencedShapeEffects(shape, theme, warnings, fillStyles, 'picture')
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
    const bookmarkNodes = children(child(mediaNode, 'bmkLst'), 'bmk')
    const mediaBookmarks: Array<{name: string; timeMs: number}> = []
    const bookmarkNames = new Set<string>()
    const bookmarkTimes = new Set<number>()
    for (const bookmark of bookmarkNodes.slice(0, 512)) {
        const name = bookmark.getAttribute('name') || ''
        const rawTime = bookmark.getAttribute('time')
        const timeMs = mediaTimeOffsetMs(rawTime)
        if (!name || name.length > 255 || rawTime === null || timeMs === undefined || bookmarkNames.has(name) || bookmarkTimes.has(timeMs)) {
            warnings.add('An invalid or duplicate embedded media bookmark was ignored.')
            continue
        }
        bookmarkNames.add(name)
        bookmarkTimes.add(timeMs)
        mediaBookmarks.push({name, timeMs})
    }
    if (bookmarkNodes.length > 512) warnings.add('Media bookmarks beyond the 512-bookmark PowerPoint limit were skipped.')
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
    const alphaInverseNodes = children(blip, 'alphaInv')
    const alphaInverseNode = alphaInverseNodes[0]
    const alphaInverseColorNode = children(alphaInverseNode)[0]
    const alphaInverseColor = animationColorChannels(alphaInverseColorNode, theme, 'rgb')
    const blipAlphaInverse = alphaInverseNode
        ? alphaInverseColor
            ? [alphaInverseColor[0] / 255, alphaInverseColor[1] / 255, alphaInverseColor[2] / 255] as [number, number, number]
            : null
        : undefined
    const imageAlphaInverse = alphaInverseNode ? blipAlphaInverse : undefined
    const imageEffectDagAlphaInverse = pictureEffects.alphaInverse
    const blipEffects = children(blip)
    const grayscaleEffectIndex = blipEffects.findIndex(effect => effect.localName === 'grayscl')
    const imageAlphaInverseAfterGrayscale = alphaInverseNode && grayscaleEffectIndex >= 0 && blipEffects.indexOf(alphaInverseNode) > grayscaleEffectIndex
        ? true
        : undefined
    if (alphaInverseNodes.length > 1) warnings.add('Only the first a:blip alpha inverse effect is rendered.')
    if (alphaInverseColorNode && !alphaInverseColor) warnings.add('An invalid a:blip alpha inverse color was ignored; alpha-only inversion was applied.')
    if (mediaUrl && (alphaInverseNode || pictureEffects.alphaInverse !== undefined)) warnings.add('Picture alpha inverse effects are not applied to video frames.')
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
        mediaBookmarks: mediaBookmarks.length ? mediaBookmarks : undefined,
        imageCrop: validCrop ? imageCrop : undefined,
        imageGrayscale: child(blip, 'grayscl') ? true : undefined,
        imageAlphaInverse: mediaUrl ? undefined : imageAlphaInverse,
        imageAlphaInverseAfterGrayscale: mediaUrl ? undefined : imageAlphaInverseAfterGrayscale,
        imageEffectDagAlphaInverse: mediaUrl ? undefined : imageEffectDagAlphaInverse,
        shadow: pictureEffects.shadow,
        shadows: pictureEffects.shadows,
        glow: pictureEffects.glow,
        glows: pictureEffects.glows,
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
        renderMatrix: [transform[0], transform[1], transform[2], transform[3], transform[4] - minX, transform[5] - minY]
    }
}

const presetConnectionGeometryCache = new Map<string, Element | null>()

async function presetConnectionGeometry(name: string): Promise<Element | undefined> {
    let cached = presetConnectionGeometryCache.get(name)
    if (cached === undefined) {
        const {default: definitions} = await import('./preset-connection-geometries')
        const source = definitions[name]
        cached = source ? parseXml(source, `The ${name} preset connection geometry`).documentElement : null
        presetConnectionGeometryCache.set(name, cached)
    }
    return cached?.cloneNode(true) as Element | undefined
}

function mergePresetConnectionAdjustments(geometry: Element, preset: Element | undefined): void {
    const defaults = child(geometry, 'avLst')
    const instance = child(preset, 'avLst')
    if (!defaults || !instance) return
    for (const adjustment of children(instance, 'gd')) {
        const name = adjustment.getAttribute('name')
        if (!name) continue
        const existing = children(defaults, 'gd').find(node => node.getAttribute('name') === name)
        if (existing) defaults.replaceChild(adjustment.cloneNode(true), existing)
        else defaults.appendChild(adjustment.cloneNode(true))
    }
}

async function resolveConnectionSite(
    reference: Element,
    targetNode: Element,
    target: PptxElement,
    warnings: Set<string>
): Promise<[number, number] | undefined> {
    const rawIndex = reference.getAttribute('idx')
    if (rawIndex === null || !/^\d+$/.test(rawIndex)) {
        warnings.add('A connector attachment has an invalid connection-site index; its saved endpoint was retained.')
        return undefined
    }
    const index = Number(rawIndex)
    const properties = child(targetNode, 'spPr')
    const custom = child(properties, 'custGeom')
    let geometry: Element | undefined
    let coordinateWidth = target.width
    let coordinateHeight = target.height
    if (custom) {
        geometry = custom
        const path = children(child(custom, 'pathLst'), 'path')[0]
        if (path) {
            coordinateWidth = numberAttr(path, 'w')
            coordinateHeight = numberAttr(path, 'h')
        }
    } else {
        const preset = child(properties, 'prstGeom')
        const presetName = preset?.getAttribute('prst') || 'rect'
        try {
            geometry = await presetConnectionGeometry(presetName)
        } catch {
            warnings.add(`The ${presetName} preset connection sites could not be loaded; the saved connector endpoint was retained.`)
            return undefined
        }
        if (geometry) mergePresetConnectionAdjustments(geometry, preset)
        if (!['sp', 'cxnSp', 'pic'].includes(targetNode.localName)) {
            warnings.add('A connector attached to a non-shape object uses rectangular connection sites.')
            try {
                geometry = await presetConnectionGeometry('rect')
            } catch {
                geometry = undefined
            }
        }
    }
    if (!(coordinateWidth > 0) || !(coordinateHeight > 0)) {
        warnings.add('A connector target has invalid connection-site coordinates; the saved endpoint was retained.')
        return undefined
    }
    const sites = children(child(geometry, 'cxnLst'), 'cxn')
    const site = sites[index]
    if (!site) {
        warnings.add(`Connection-site index ${index} is unavailable on target ${target.id}; the saved connector endpoint was retained.`)
        return undefined
    }
    const position = child(site, 'pos')
    const xToken = position?.getAttribute('x')
    const yToken = position?.getAttribute('y')
    if (xToken === null || xToken === undefined || yToken === null || yToken === undefined) {
        warnings.add(`Connection site ${index} on target ${target.id} has no position; the saved connector endpoint was retained.`)
        return undefined
    }
    const resolveGuide = customGeometryGuideResolver(geometry, coordinateWidth, coordinateHeight, warnings)
    const x = resolveGuide(xToken)
    const y = resolveGuide(yToken)
    if (x === undefined || y === undefined) {
        warnings.add(`Connection site ${index} on target ${target.id} uses an unresolved guide; the saved connector endpoint was retained.`)
        return undefined
    }
    return [x / coordinateWidth * target.width, y / coordinateHeight * target.height]
}

function localPointToSlide(element: PptxElement, [x, y]: [number, number]): [number, number] {
    if (element.renderMatrix) {
        const [a, b, c, d, e, f] = element.renderMatrix
        return [element.x + a * x + c * y + e, element.y + b * x + d * y + f]
    }
    const localX = element.flipH ? element.width - x : x
    const localY = element.flipV ? element.height - y : y
    const radians = element.rotation * Math.PI / 180
    const cosine = Math.cos(radians)
    const sine = Math.sin(radians)
    const dx = localX - element.width / 2
    const dy = localY - element.height / 2
    return [element.x + element.width / 2 + dx * cosine - dy * sine, element.y + element.height / 2 + dx * sine + dy * cosine]
}

function slidePointToLocal(element: PptxElement, [x, y]: [number, number]): [number, number] | undefined {
    if (element.renderMatrix) {
        const [a, b, c, d, e, f] = element.renderMatrix
        const determinant = a * d - b * c
        if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) return undefined
        const dx = x - element.x - e
        const dy = y - element.y - f
        return [(d * dx - c * dy) / determinant, (-b * dx + a * dy) / determinant]
    }
    const radians = element.rotation * Math.PI / 180
    const cosine = Math.cos(radians)
    const sine = Math.sin(radians)
    const dx = x - element.x - element.width / 2
    const dy = y - element.y - element.height / 2
    const localX = element.width / 2 + dx * cosine + dy * sine
    const localY = element.height / 2 - dx * sine + dy * cosine
    return [element.flipH ? element.width - localX : localX, element.flipV ? element.height - localY : localY]
}

function connectorPathEndpoints(path: string): {start: [number, number]; end: [number, number]} | undefined {
    const numberPattern = /[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi
    const matches = [...path.matchAll(/([MLC])([^MLC]*)/g)]
    if (!matches.length || /[^\s,]/.test(path.replace(/[MLC]|[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi, ''))) return undefined
    let start: [number, number] | undefined
    let end: [number, number] | undefined
    let moveCount = 0
    for (const match of matches) {
        const command = match[1]!
        const values = [...match[2]!.matchAll(numberPattern)].map(value => Number(value[0]))
        if (!values.length || values.some(value => !Number.isFinite(value)) || values.length % (command === 'C' ? 6 : 2) !== 0) return undefined
        if (command === 'M') {
            if (start) return undefined
            start = [values[0]!, values[1]!]
            moveCount++
        } else if (!start) return undefined
        end = [values.at(-2)!, values.at(-1)!]
    }
    return moveCount === 1 && start && end ? {start, end} : undefined
}

function alignConnectorPath(path: string, start: [number, number], end: [number, number]): string | undefined {
    const endpoints = connectorPathEndpoints(path)
    if (!endpoints) return undefined
    const [x1, y1] = endpoints.start
    const [x2, y2] = endpoints.end
    const dx = x2 - x1
    const dy = y2 - y1
    const targetDx = end[0] - start[0]
    const targetDy = end[1] - start[1]
    let map: (x: number, y: number) => [number, number]
    if (Math.abs(dx) > 1e-8 && Math.abs(dy) > 1e-8) {
        const scaleX = targetDx / dx
        const scaleY = targetDy / dy
        map = (x, y) => [start[0] + (x - x1) * scaleX, start[1] + (y - y1) * scaleY]
    } else {
        const sourceLength = Math.hypot(dx, dy)
        if (sourceLength < 1e-8) return undefined
        const targetLength = Math.hypot(targetDx, targetDy)
        const scale = targetLength / sourceLength
        const cosine = targetLength === 0 ? 1 : (dx * targetDx + dy * targetDy) / (sourceLength * targetLength)
        const sine = targetLength === 0 ? 0 : (dx * targetDy - dy * targetDx) / (sourceLength * targetLength)
        map = (x, y) => {
            const localX = x - x1
            const localY = y - y1
            return [start[0] + scale * (cosine * localX - sine * localY), start[1] + scale * (sine * localX + cosine * localY)]
        }
    }
    const numberPattern = /[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi
    const formatted = (value: number) => {
        const rounded = Number(value.toFixed(4))
        return Object.is(rounded, -0) ? '0' : String(rounded)
    }
    return [...path.matchAll(/([MLC])([^MLC]*)/g)].map(match => {
        const command = match[1]!
        const values = [...match[2]!.matchAll(numberPattern)].map(value => Number(value[0]))
        const mapped: string[] = []
        for (let index = 0; index < values.length; index += 2) {
            const point = map(values[index]!, values[index + 1]!)
            mapped.push(formatted(point[0]), formatted(point[1]))
        }
        return `${command} ${mapped.join(' ')}`
    }).join(' ')
}

async function resolveConnectorAttachments(tree: Element | undefined, elements: PptxElement[], warnings: Set<string>): Promise<void> {
    const elementById = new Map(elements.map(element => [element.id, element]))
    const nodeById = new Map<string, Element>()
    for (const identity of descendants(tree, 'cNvPr')) {
        const id = identity.getAttribute('id')
        let node = identity.parentElement
        while (node && !['sp', 'cxnSp', 'pic', 'graphicFrame', 'grpSp'].includes(node.localName)) node = node.parentElement
        if (id && node) nodeById.set(id, node)
    }
    for (const connectorNode of nodeById.values()) {
        if (connectorNode.localName !== 'cxnSp') continue
        const connectionProperties = child(child(connectorNode, 'nvCxnSpPr'), 'cNvCxnSpPr')
        const startReference = child(connectionProperties, 'stCxn')
        const endReference = child(connectionProperties, 'endCxn')
        if (!startReference && !endReference) continue
        const connectorId = firstDescendant(connectorNode, 'cNvPr')?.getAttribute('id') || ''
        const connector = elementById.get(connectorId)
        const route = connector?.linePath || (connector?.geometry === 'lineInv'
            ? 'M 0 1000 L 1000 0'
            : connector && ['line', 'straightConnector1'].includes(connector.geometry) ? 'M 0 0 L 1000 1000' : undefined)
        if (!connector || !route || !(connector.width > 0) || !(connector.height > 0)) {
            warnings.add(`Connector ${connectorId || '(unknown)'} has no supported route; its saved endpoint positions were retained.`)
            continue
        }
        const endpoints = connectorPathEndpoints(route)
        if (!endpoints) {
            warnings.add(`Connector ${connectorId} uses an unsupported route; its saved endpoint positions were retained.`)
            continue
        }
        const resolveEndpoint = async (reference: Element | undefined, point: [number, number]): Promise<[number, number]> => {
            if (!reference) return point
            const targetId = reference.getAttribute('id')
            const target = targetId ? elementById.get(targetId) : undefined
            const targetNode = targetId ? nodeById.get(targetId) : undefined
            if (!target || !targetNode) {
                warnings.add(`Connector ${connectorId} references missing target ${targetId || '(unknown)'}; its saved endpoint was retained.`)
                return point
            }
            const site = await resolveConnectionSite(reference, targetNode, target, warnings)
            if (!site) return point
            const slidePoint = localPointToSlide(target, site)
            const localPoint = slidePointToLocal(connector, slidePoint)
            if (!localPoint || !localPoint.every(Number.isFinite)) {
                warnings.add(`Connector ${connectorId} has a non-invertible transform; its saved endpoint was retained.`)
                return point
            }
            return [localPoint[0] / connector.width * 1000, localPoint[1] / connector.height * 1000]
        }
        const start = await resolveEndpoint(startReference, endpoints.start)
        const end = await resolveEndpoint(endReference, endpoints.end)
        const aligned = alignConnectorPath(route, start, end)
        if (!aligned) {
            warnings.add(`Connector ${connectorId} could not align its route to the requested connection sites; its saved route was retained.`)
            continue
        }
        connector.linePath = aligned
        if (!['line', 'lineInv', 'straightConnector1'].includes(connector.geometry)) {
            warnings.add(`Connector ${connectorId} retains its preset route with endpoint warping; PowerPoint automatic rerouting and connection-site angles are approximated.`)
        }
    }
}

async function parseSceneElements(node: Element, rels: Map<string, PptxRelationship>, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>, parentTransform: Matrix2D | undefined, frameAnimationTargetIds: Set<string>, sourceTheme: Record<string, string>, fillStyles?: ParsedTheme, groupFill?: string, smartArtAnimationTargets = new Map<string, string>()): Promise<PptxElement[]> {
    if (node.localName === 'grpSp') {
        const localTransform = groupTransformMatrix(node)
        if (!localTransform) {
            warnings.add('A grouped shape with invalid coordinate extents was skipped.')
            return []
        }
        const transform = parentTransform ? multiplyMatrix(parentTransform, localTransform) : localTransform
        const groupFillNode = children(child(node, 'grpSpPr')).find(item => FILL_KINDS.includes(item.localName))
        const currentGroupFill = groupFillNode ? fillColor(groupFillNode, theme, warnings, groupFill, true) : undefined
        const elements: PptxElement[] = []
        for (const item of children(node)) {
            if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(item.localName)) continue
            elements.push(...(await parseSceneElements(item, rels, zip, theme, warnings, urls, imageUrlCache, transform, frameAnimationTargetIds, sourceTheme, fillStyles, currentGroupFill, smartArtAnimationTargets)))
        }
        return elements
    }

    let element: PptxElement | undefined
    if (node.localName === 'sp' || node.localName === 'cxnSp') element = parseShape(node, theme, warnings, {}, {}, true, rels, fillStyles, groupFill)
    else if (node.localName === 'pic') element = await parsePicture(node, rels, zip, theme, warnings, urls, imageUrlCache, fillStyles)
    else if (node.localName === 'graphicFrame') {
        element = parseTable(node, theme, warnings, rels)
        if (!element && firstDescendant(node, 'chart')) element = await parseChart(node, rels, zip, theme, warnings, sourceTheme)
        else if (!element) {
            element = await parseOlePreview(node, rels, zip, theme, warnings, urls, imageUrlCache, fillStyles)
            if (!element) {
                const cachedDrawing = await parseSmartArtCachedDrawing(node, rels, zip, theme, warnings, urls, imageUrlCache, frameAnimationTargetIds, parentTransform, sourceTheme, fillStyles, smartArtAnimationTargets)
                if (cachedDrawing) return cachedDrawing
                element = await parseSmartArtTextFallback(node, rels, zip, theme, warnings)
            }
            if (!element) warnings.add('Unsupported graphic frames are not rendered yet.')
        }
    }
    if (element && node.localName === 'cxnSp' && !LINE_GEOMETRIES.has(element.geometry)) warnings.add('Connector geometry is not fully rendered yet.')
    if (!element) return []
    if (!element.hyperlink) element.hyperlink = parseHyperlink(child(firstDescendant(node, 'cNvPr'), 'hlinkClick'), rels, warnings)
    return [parentTransform ? applyGroupTransform(element, parentTransform) : element]
}

async function backgroundFill(xml: XMLDocument | undefined, rels: Map<string, PptxRelationship>, theme: ParsedTheme, warnings: Set<string>, zip: JSZip, urls: string[], imageUrlCache: Map<string, string>): Promise<{color?: string; image?: PptxSlide['backgroundImage']}> {
    const background = xml && firstDescendant(xml.documentElement, 'bg')
    const properties = child(background, 'bgPr')
    if (!properties) {
        const reference = child(background, 'bgRef')
        if (reference) {
            // idx 1–999 selects fmtScheme.fillStyleLst; idx 1001+ selects bgFillStyleLst.
            // 0 and 1000 mean no background style. See Open XML BackgroundStyleReference.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.background.backgroundstylereference?view=openxml-3.0.1
            const rawIndex = reference.getAttribute('idx')?.trim() || ''
            const index = /^\d+$/.test(rawIndex) ? Number(rawIndex) : Number.NaN
            if (!Number.isSafeInteger(index) || index < 0) {
                warnings.add('A theme slide background reference with an invalid style index was ignored.')
                return {}
            }
            if (index === 0 || index === 1000) return {}
            const styles = index < 1000 ? theme.fillStyles : theme.backgroundFillStyles
            const style = styles[index < 1000 ? index - 1 : index - 1001]
            if (!style) {
                warnings.add(`Theme slide background style ${index} could not be resolved.`)
                return {}
            }
            const colorReference = children(reference).find(node => ['srgbClr', 'scrgbClr', 'sysClr', 'schemeClr', 'hslClr', 'prstClr'].includes(node.localName))
            const placeholderColor = readColor(colorReference, theme.colors)
            const styleTheme = placeholderColor ? {...theme.colors, phClr: placeholderColor} : theme.colors
            const color = fillColor(style, styleTheme, warnings, undefined, true)
            return {color: color === 'transparent' ? undefined : color}
        }
        return {}
    }
    const imageFill = child(properties, 'blipFill')
    if (!imageFill) {
        if (child(properties, 'grpFill')) warnings.add('Group slide background fills are not rendered.')
        const color = fillColor(properties, theme.colors, warnings, undefined, true)
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
        if (flip !== rawFlip) warnings.add(`Unknown slide background tile flip "${rawFlip}" was replaced with none.`)
        return {
            image: {
                url,
                crop,
                tile: {scaleX, scaleY, align, flip, offsetX: readTileOffset('tx'), offsetY: readTileOffset('ty')}
            }
        }
    }
    return {image: {url, crop}}
}

async function parseDecorations(path: string, xml: XMLDocument, zip: JSZip, theme: Record<string, string>, warnings: Set<string>, urls: string[], imageUrlCache: Map<string, string>, sourceTheme: Record<string, string>, fillStyles?: ParsedTheme): Promise<PptxElement[]> {
    const tree = firstDescendant(xml.documentElement, 'spTree')
    const rels = await loadRelationships(zip, relationshipPath(path))
    const elements: PptxElement[] = []
    for (const node of children(tree)) {
        if (node.localName === 'sp' && firstDescendant(node, 'ph')) continue
        elements.push(...(await parseSceneElements(node, rels, zip, theme, warnings, urls, imageUrlCache, undefined, new Set(), sourceTheme, fillStyles)))
    }
    return elements.map(element => ({...element, id: `${path}:${element.id}`}))
}

export async function parsePptx(file: File): Promise<PptxDocument> {
    await validatePptxFile(file)
    const urls: string[] = []
    try {
        const zip = await JSZip.loadAsync(file)
        await validateArchiveSize(zip)
        const presentationPath = 'ppt/presentation.xml'
        const presentationEntry = zip.file(presentationPath)
        if (!presentationEntry) throw new Error('This file is not a valid PowerPoint presentation.')
        const presentation = parseXml(await presentationEntry.async('string'), presentationPath)
        const presentationRels = await loadRelationships(zip, 'ppt/_rels/presentation.xml.rels')
        const size = firstDescendant(presentation.documentElement, 'sldSz')
        const slideSizeCoordinate = (attribute: 'cx' | 'cy', fallback: number): number => {
            const raw = size?.getAttribute(attribute)
            if (raw === null || raw === undefined) return fallback
            const value = Number(raw)
            if (!Number.isSafeInteger(value) || value < 914_400 || value > 51_206_400) {
                throw new Error('ppt/presentation.xml contains an invalid slide size.')
            }
            return value
        }
        const width = slideSizeCoordinate('cx', 12_192_000)
        const height = slideSizeCoordinate('cy', 6_858_000)
        const slideIds = firstDescendant(presentation.documentElement, 'sldIdLst')
        const orderedSlides = children(slideIds, 'sldId')
        if (!orderedSlides.length) throw new Error('No slides were found in this presentation.')
        if (orderedSlides.length > 500) throw new Error('This presentation contains more than 500 slides.')
        const presetShapeDefinitionsReady = loadPresetShapeDefinitions()
        const fallbackThemePath = Object.keys(zip.files).find(path => /^ppt\/theme\/theme\d+\.xml$/i.test(path))
        const warnings = new Set<string>()
        if (!(await presetShapeDefinitionsReady)) warnings.add('The bundled ECMA-376 preset shape definitions could not be loaded; unhandled preset shapes will use rectangle fallbacks.')
        const slides: PptxSlide[] = []
        const imageUrlCache = new Map<string, string>()
        const themeCache = new Map<string, ParsedTheme>()
        const decorationCache = new Map<string, Promise<PptxElement[]>>()
        const embeddedFonts = await parseEmbeddedFonts(presentation, presentationRels, zip, urls, warnings)

        const readXml = async (path: string | undefined): Promise<XMLDocument | undefined> => {
            if (!path) return undefined
            const entry = zip.file(path)
            return entry ? parseXml(await entry.async('string'), path) : undefined
        }
        const appPropertiesXml = await readXml('docProps/app.xml')
        const sourceApplication = firstDescendant(appPropertiesXml?.documentElement, 'Application')?.textContent?.trim() || ''
        const isWpsApplication = /wps|kingsoft/i.test(sourceApplication)
        const presentationPropertiesRelation = [...presentationRels.values()].find(relation => relation.type.endsWith('/presProps'))
        const presentationPropertiesXml = presentationPropertiesRelation && !presentationPropertiesRelation.external
            ? await readXml(presentationPropertiesRelation.target)
            : undefined
        if (presentationPropertiesRelation?.external) warnings.add('The external presentation-properties relationship was ignored.')
        else if (presentationPropertiesRelation && !presentationPropertiesXml) warnings.add('The presentation references missing presentation properties.')
        const showSettings = parsePresentationShowSettings(child(presentationPropertiesXml?.documentElement, 'showPr'), warnings)
        const loadTheme = async (masterRels: Map<string, PptxRelationship>): Promise<ParsedTheme> => {
            const themePath = [...masterRels.values()].find(relation => relation.type.endsWith('/theme'))?.target || fallbackThemePath || 'default'
            const cached = themeCache.get(themePath)
            if (cached) return cached
            const themeXml = await readXml(themePath === 'default' ? undefined : themePath)
            const theme = parseTheme(themeXml)
            themeCache.set(themePath, theme)
            return theme
        }
        const loadDecorations = (path: string | undefined, xml: XMLDocument | undefined, theme: Record<string, string>, sourceTheme: Record<string, string>, fillStyles?: ParsedTheme): Promise<PptxElement[]> => {
            if (!path || !xml) return Promise.resolve([])
            const cacheKey = `${path}|${Object.entries(theme).concat(Object.entries(sourceTheme))
                .map(([key, value]) => `${key}:${value}`)
                .join(';')}`
            let cached = decorationCache.get(cacheKey)
            if (!cached) {
                cached = parseDecorations(path, xml, zip, theme, warnings, urls, imageUrlCache, sourceTheme, fillStyles)
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
            const notesRels = notesXml && notesRelation ? await loadRelationships(zip, relationshipPath(notesRelation.target)) : new Map<string, PptxRelationship>()
            const notesMasterRelation = [...notesRels.values()].find(relation => relation.type.endsWith('/notesMaster'))
            const notesMasterXml = notesMasterRelation && !notesMasterRelation.external ? await readXml(notesMasterRelation.target) : undefined
            if (notesMasterRelation?.external) warnings.add(`Slide ${index + 1} has an external notes-master relationship that was ignored.`)
            else if (notesMasterRelation && !notesMasterXml) warnings.add(`Slide ${index + 1} references a missing notes master.`)
            const notesMasterRels = notesMasterXml && notesMasterRelation ? await loadRelationships(zip, relationshipPath(notesMasterRelation.target)) : new Map<string, PptxRelationship>()
            const notesThemeData = notesXml ? await loadTheme(notesMasterRels) : undefined
            const notesMasterMap = resolveColorMap(notesMasterXml?.documentElement, DEFAULT_COLOR_MAP, DEFAULT_COLOR_MAP, warnings)
            const notesMap = resolveColorMap(notesXml?.documentElement, notesMasterMap, notesMasterMap, warnings)
            const notesTheme = mappedThemeColors(notesThemeData?.colors || DEFAULT_THEME, notesMap)
            const layoutPath = [...rels.values()].find(relation => relation.type.endsWith('/slideLayout'))?.target
            const layoutXml = await readXml(layoutPath)
            if (layoutPath && !layoutXml) warnings.add(`Slide ${index + 1} references a missing layout.`)
            const layoutRels = layoutPath ? await loadRelationships(zip, relationshipPath(layoutPath)) : new Map<string, PptxRelationship>()
            const masterPath = [...layoutRels.values()].find(relation => relation.type.endsWith('/slideMaster'))?.target
            const masterXml = await readXml(masterPath)
            if (masterPath && !masterXml) warnings.add(`Slide ${index + 1} references a missing master.`)
            const masterRels = masterPath ? await loadRelationships(zip, relationshipPath(masterPath)) : new Map<string, PptxRelationship>()
            const themeData = await loadTheme(masterRels)
            const masterColorMap = resolveColorMap(masterXml?.documentElement, DEFAULT_COLOR_MAP, DEFAULT_COLOR_MAP, warnings)
            const layoutColorMap = resolveColorMap(layoutXml?.documentElement, masterColorMap, masterColorMap, warnings)
            const slideColorMap = resolveColorMap(xml.documentElement, layoutColorMap, masterColorMap, warnings)
            const masterTheme = mappedThemeColors(themeData.colors, masterColorMap)
            const layoutTheme = mappedThemeColors(themeData.colors, layoutColorMap)
            const theme = mappedThemeColors(themeData.colors, slideColorMap)
            const masterTextStyles = parseMasterTextStyles(masterXml, theme, warnings, masterRels)
            const masterPlaceholders = parsePlaceholderCandidates(masterXml, theme, warnings, masterTextStyles, [], false, masterRels, themeData)
            const layoutPlaceholders = parsePlaceholderCandidates(layoutXml, theme, warnings, masterTextStyles, masterPlaceholders, false, layoutRels, themeData)
            const tree = firstDescendant(xml.documentElement, 'spTree')
            const frameAnimationTargetIds = new Set(descendants(xml.documentElement, 'spTgt').filter(target => !child(target, 'subSp')).map(target => target.getAttribute('spid') || '').filter(Boolean))
            const smartArtAnimationTargets = new Map<string, string>()
            // p:sld and p:sldLayout showMasterSp both default to true and gate master-slide shapes.
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slide.showmastershapes?view=openxml-3.0.1
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slidelayout.showmastershapes?view=openxml-3.0.1
            const masterShapesVisible = [xml.documentElement, layoutXml?.documentElement].every(root => {
                const value = root?.getAttribute('showMasterSp')
                return value === null || value === undefined || !['0', 'false'].includes(value.trim().toLowerCase())
            })
            const masterElements = masterShapesVisible ? await loadDecorations(masterPath, masterXml, masterTheme, themeData.colors, themeData) : []
            const layoutElements = await loadDecorations(layoutPath, layoutXml, layoutTheme, themeData.colors, themeData)
            const elements: PptxElement[] = [...masterElements, ...layoutElements]
            const slideElementStart = elements.length
            for (const node of children(tree)) {
                if (node.localName === 'sp') {
                    const identity = parsePlaceholderIdentity(node)
                    const layoutPlaceholder = matchPlaceholder(identity, layoutPlaceholders)
                    const masterPlaceholder = matchPlaceholder(identity, masterPlaceholders)
                    if (!identity) {
                        elements.push(parseShape(node, theme, warnings, {}, {}, true, rels, themeData))
                        continue
                    }
                    const masterStyle = masterTextStyles[placeholderStyleCategory(identity.type)]
                    const inheritedParagraphStyles = layoutPlaceholder?.paragraphStyles || masterPlaceholder?.paragraphStyles || {}
                    let element = parseShape(node, theme, warnings, masterStyle?.style || {}, inheritedParagraphStyles, true, rels, themeData)
                    element = inheritPlaceholderProperties(element, node, layoutPlaceholder || masterPlaceholder, warnings)
                    elements.push(element)
                } else if (['pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(node.localName)) {
                    elements.push(...(await parseSceneElements(node, rels, zip, theme, warnings, urls, imageUrlCache, undefined, frameAnimationTargetIds, themeData.colors, themeData, undefined, smartArtAnimationTargets)))
                }
            }
            await resolveConnectorAttachments(tree, elements.slice(slideElementStart), warnings)
            await resolvePictureBullets(elements, zip, warnings, urls, imageUrlCache)
            const notes = parseSpeakerNotes(notesXml, notesMasterXml, notesTheme, warnings, notesRels, notesMasterRels)
            if (notes.speakerNoteParagraphs) await resolvePictureBullets([{paragraphs: notes.speakerNoteParagraphs}], zip, warnings, urls, imageUrlCache)
            const animations = await parseAnimations(xml.documentElement, slidePath, theme, warnings, elements, width, height, smartArtAnimationTargets, rels, zip, urls, imageUrlCache, isWpsApplication)
            const slideInfo = firstDescendant(xml.documentElement, 'cSld')
            const slideBackground = await backgroundFill(xml, rels, {...themeData, colors: theme}, warnings, zip, urls, imageUrlCache)
            const layoutBackground = await backgroundFill(layoutXml, layoutRels, {...themeData, colors: layoutTheme}, warnings, zip, urls, imageUrlCache)
            const masterBackground = await backgroundFill(masterXml, masterRels, {...themeData, colors: masterTheme}, warnings, zip, urls, imageUrlCache)
            const selectedBackground = slideBackground.image || slideBackground.color ? slideBackground : layoutBackground.image || layoutBackground.color ? layoutBackground : masterBackground
            slides.push({
                id: relId || `slide-${index + 1}`,
                packagePath: slidePath,
                name: child(slideInfo, 'name')?.getAttribute('val') || `Slide ${index + 1}`,
                hidden: parseSlideHidden(xml.documentElement, warnings),
                elements,
                ...notes,
                ...animations,
                ...parseSlideAdvance(xml.documentElement, warnings, layoutXml?.documentElement, masterXml?.documentElement),
                background: selectedBackground.color || '#ffffff',
                backgroundImage: selectedBackground.image,
                transition: await parseTransition(xml.documentElement, warnings, rels, zip, urls, imageUrlCache, layoutXml?.documentElement, layoutRels, masterXml?.documentElement, masterRels)
            })
        }
        if (!slides.length) throw new Error('No readable slides were found in this presentation.')
        const slidePaths = new Set(slides.map(slide => slide.packagePath))
        for (const slide of slides) {
            for (const element of slide.elements) {
                const hyperlinks = [
                    element.hyperlink,
                    ...element.paragraphs.flatMap(paragraph => paragraph.runs.map(run => run.hyperlink)),
                    ...(element.table?.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs.flatMap(paragraph => paragraph.runs.map(run => run.hyperlink)))) || [])
                ]
                if (hyperlinks.some(link => link?.kind === 'slide' && !slidePaths.has(link.targetPath))) {
                    warnings.add('An internal slide hyperlink points to a missing slide and was ignored.')
                }
            }
        }
        return {name: file.name, width, height, slides, showSettings, warnings: [...warnings], objectUrls: urls, embeddedFonts}
    } catch (error) {
        urls.forEach(URL.revokeObjectURL)
        throw normalizePptxError(error)
    }
}
