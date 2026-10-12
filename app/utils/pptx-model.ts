import type { TimingCurve } from './animation-timing'

export interface PptxRun {
    text: string
    language?: string
    fontFamily?: string
    fontFamilyEastAsia?: string
    fontFamilyComplexScript?: string
    fontSizePt?: number
    letterSpacingPt?: number
    baselineOffsetEm?: number
    color?: string
    fillPattern?: string
    fillGradient?: string
    highlightColor?: string
    outlineColor?: string
    outlineWidth?: number
    bold?: boolean
    italic?: boolean
    underline?: boolean
    underlineStyle?: 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy'
    underlineColor?: string
    underlineThicknessEmu?: number
    lineThrough?: boolean
    lineThroughStyle?: 'solid' | 'double'
    hyperlink?: PptxHyperlink
    shadow?: PptxShadow
    shadows?: PptxShadow[]
    glow?: PptxGlow
    glows?: PptxGlow[]
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
    inset?: boolean
}

export interface PptxGlow {
    color: string
    radius: number
}

export interface PptxParagraph {
    align: 'left' | 'center' | 'right' | 'justify'
    direction?: 'ltr' | 'rtl'
    tabStopInterval?: number
    runs: PptxRun[]
    level?: number
    marginLeft?: number
    marginRight?: number
    indent?: number
    spaceBefore?: {unit: 'emu' | 'ratio'; value: number}
    spaceAfter?: {unit: 'emu' | 'ratio'; value: number}
    lineSpacing?: {unit: 'emu' | 'ratio'; value: number}
    bullet?: ({kind: 'char'; char: string} | {kind: 'number'; type: string; startAt?: number} | {kind: 'image'; target?: string; url?: string}) & {size?: {unit: 'ratio' | 'points'; value: number}}
}

export interface PptxTableCell {
    paragraphs: PptxParagraph[]
    textAnchorCenter?: boolean
    textColumnCount?: number
    textColumnSpacing?: number
    textColumnsRightToLeft?: boolean
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

export type PptxChartDataLabelStyle = Pick<PptxRun, 'fontFamily' | 'fontSizePt' | 'color' | 'bold' | 'italic' | 'letterSpacingPt' | 'baselineOffsetEm' | 'underline' | 'underlineStyle' | 'underlineColor' | 'underlineThicknessEmu' | 'lineThrough' | 'lineThroughStyle'>

export type PptxChartMarkerSymbol = 'auto' | 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'picture' | 'plus' | 'square' | 'star' | 'triangle' | 'x'

export interface PptxChartSeries {
    legendIndex: number
    name: string
    values: Array<number | null>
    chartType?: 'bar' | 'line' | 'area' | 'stock'
    chartGroupIndex?: number
    grouping?: PptxChartGrouping
    valueAxisId?: string
    lineShowLine?: boolean
    lineShowMarker?: boolean
    lineSmooth?: boolean
    xValues?: Array<number | null>
    bubbleSizes?: Array<number | null>
    scatterShowLine?: boolean
    scatterShowMarker?: boolean
    scatterSmooth?: boolean
    radarShowLine?: boolean
    radarShowMarker?: boolean
    radarMarkerSize?: number
    markerSymbol?: PptxChartMarkerSymbol
    markerSize?: number
    markerSymbolOverrides?: Record<number, PptxChartMarkerSymbol>
    markerSizeOverrides?: Record<number, number>
    color: string
    stroke?: string
    strokeWidthEmu?: number
    strokeDash?: PptxElement['strokeDash']
    strokeCap?: PptxElement['strokeCap']
    pointColors: Record<number, string>
    pointStrokes?: Record<number, string>
    pointStrokeWidthsEmu?: Record<number, number>
    pointStrokeDashes?: Record<number, NonNullable<PptxElement['strokeDash']>>
    pointStrokeCaps?: Record<number, NonNullable<PptxElement['strokeCap']>>
    showLegendKey: boolean
    legendKeyOverrides: Record<number, boolean>
    showValueLabels: boolean
    valueLabelOverrides: Record<number, boolean>
    showPercentLabels: boolean
    percentLabelOverrides: Record<number, boolean>
    showBubbleSizeLabels: boolean
    bubbleSizeLabelOverrides: Record<number, boolean>
    showCategoryNameLabels: boolean
    categoryNameLabelOverrides: Record<number, boolean>
    showSeriesNameLabels: boolean
    seriesNameLabelOverrides: Record<number, boolean>
    labelSeparator?: string
    labelSeparatorOverrides: Record<number, string>
    numberFormat?: string
    cachedNumberFormat?: string
    xNumberFormat?: string
    valueNumberFormatOverrides: Record<number, string>
    dataLabelPosition?: PptxChartDataLabelPosition
    dataLabelPositionOverrides: Record<number, PptxChartDataLabelPosition>
    showLeaderLines?: boolean
    leaderLineOverrides?: Record<number, boolean>
    dataLabelStyle?: PptxChartDataLabelStyle
    dataLabelStyleOverrides: Record<number, PptxChartDataLabelStyle>
}

export type PptxChartGrouping = 'clustered' | 'stacked' | 'percentStacked'

export interface PptxChartAxis {
    min: number
    max: number
    logBase?: number
    majorUnit?: number
    minorUnit?: number
    numberFormat?: string
    reverse?: boolean
    deleted?: boolean
    majorGridlines?: boolean
    position?: PptxChartAxisPosition
    crosses?: PptxChartAxisCrosses
    crossesAt?: number
    majorTickMark?: PptxChartTickMark
    minorTickMark?: PptxChartTickMark
    tickLabelPosition?: PptxChartTickLabelPosition
}

export interface PptxChartStackSegment {
    start: number
    end: number
}

export interface PptxChart {
    type: 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar' | 'stock'
    direction?: 'horizontal' | 'vertical'
    grouping?: PptxChartGrouping
    radarStyle?: 'standard' | 'marker' | 'filled'
    stock?: {
        highIndex: number
        lowIndex: number
        closeIndex: number
        openIndex?: number
        volumeIndex?: number
        volumeGapWidth?: number
        highLowColor: string
        showHighLowLines: boolean
        showDropLines: boolean
        showUpDownBars: boolean
        upColor: string
        downColor: string
        gapWidth: number
    }
    holeSize?: number
    firstSliceAngle?: number
    bubbleScale?: number
    bubbleSizeRepresents?: 'area' | 'w'
    showNegativeBubbles?: boolean
    bubbleVaryColors?: boolean
    title?: string
    categories: string[]
    palette: string[]
    indexedColors?: string[]
    series: PptxChartSeries[]
    valueAxes?: Array<{id: string; axis: PptxChartAxis; primary?: boolean}>
    labelSeparator?: string
    numberFormat?: string
    date1904?: boolean
    dateAxis?: {values: number[]; min?: number; max?: number; reverse?: boolean; majorUnit?: number; majorTimeUnit?: 'days' | 'months' | 'years'; minorUnit?: number; minorTimeUnit?: 'days' | 'months' | 'years'}
    valueAxis?: PptxChartAxis
    xAxis?: PptxChartAxis
    categoryAxisNumberFormat?: string
    categoryAxisPosition?: PptxChartAxisPosition
    categoryAxisCrosses?: PptxChartAxisCrosses
    categoryAxisCrossesAt?: number
    categoryAxisDeleted?: boolean
    categoryAxisMajorGridlines?: boolean
    categoryAxisMajorTickMark?: PptxChartTickMark
    categoryAxisMinorTickMark?: PptxChartTickMark
    categoryAxisTickLabelPosition?: PptxChartTickLabelPosition
    categoryAxisTickLabelSkip?: number
    categoryAxisTickMarkSkip?: number
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
    fillMode?: 'lighten' | 'lightenLess' | 'darken' | 'darkenLess'
    stroke: boolean
}

export type PptxTextOrientation = 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'

export type PptxLineEnd = {type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow' | 'chevron' | 'doubleChevron'; width: 'sm' | 'med' | 'lg'; length: 'sm' | 'med' | 'lg'}

export interface PptxElement {
    id: string
    name: string
    hyperlink?: PptxHyperlink
    kind: 'shape' | 'picture' | 'video' | 'audio' | 'line' | 'placeholder' | 'table' | 'chart'
    geometry: string
    linePath?: string
    presetAdjustment?: number
    presetAdjustments?: Record<string, number>
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
    strokePattern?: string
    strokeWidth: number
    strokeDash?: string | number[]
    strokeCap?: 'butt' | 'round' | 'square'
    lineHeadEnd?: PptxLineEnd
    lineTailEnd?: PptxLineEnd
    shadow?: PptxShadow
    shadows?: PptxShadow[]
    glow?: PptxGlow
    glows?: PptxGlow[]
    imageUrl?: string
    mediaUrl?: string
    mediaType?: 'video' | 'audio'
    mediaMime?: string
    mediaFormat?: string
    mediaTrim?: {startMs: number; endMs: number}
    mediaFade?: {inMs: number; outMs: number}
    mediaBookmarks?: Array<{name: string; timeMs: number}>
    imageCrop?: {left: number; top: number; right: number; bottom: number}
    imageGrayscale?: boolean
    imageAlphaInverse?: [number, number, number] | null
    imageAlphaInverseAfterGrayscale?: boolean
    imageEffectDagAlphaInverse?: [number, number, number] | null
    table?: PptxTable
    chart?: PptxChart
    paragraphs: PptxParagraph[]
    margins: {left: number; right: number; top: number; bottom: number}
    verticalAlign: 'top' | 'middle' | 'bottom'
    textWrap?: 'square' | 'none'
    textHorizontalOverflow?: 'overflow' | 'clip'
    textVerticalOverflow?: 'overflow' | 'clip' | 'ellipsis'
    textOrientation?: PptxTextOrientation
    textAnchorCenter?: boolean
    textColumnCount?: number
    textColumnSpacing?: number
    textColumnsRightToLeft?: boolean
    textFontScale?: number
    textLineSpacingReduction?: number
    textAutoFit?: 'normal' | 'shape'
    textAutoFitDynamic?: boolean
    placeholder?: {type: string; index?: string}
}

export type PptxAnimationTriggerCondition =
    | {type: 'shape'; id: string; event: 'onClick' | 'onDblClick' | 'onMouseOver' | 'onMouseOut'; delayMs: number}
    | {type: 'slide'; event: 'onNext' | 'onPrev' | 'onStopAudio'; delayMs: number}
    | {type: 'mediaBookmark'; id: string; bookmarkName: string; timeMs: number; delayMs: number}
    | {type: 'timeNode'; id: string; event: 'begin' | 'onBegin' | 'end' | 'onEnd'; sourceStep?: number; sourceDelayMs: number; sourceMediaActionKey?: string; delayMs: number}

export interface PptxSlide {
    id: string
    packagePath?: string
    name: string
    hidden?: boolean
    elements: PptxElement[]
    background: string
    backgroundImage?: {
        url: string
        crop: {left: number; top: number; right: number; bottom: number}
        tile?: {scaleX: number; scaleY: number; align: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br'; flip: 'none' | 'x' | 'y' | 'xy'; offsetX: number; offsetY: number}
    }
    speakerNotes?: string
    speakerNoteParagraphs?: PptxParagraph[]
    automaticAnimations?: PptxAnimation[]
    animationSteps?: PptxAnimation[][]
    animationSequence?: {
        concurrent?: boolean
        nextAction: 'none' | 'seek'
        advancesOnNext?: boolean
        rewindsOnPrevious?: boolean
    }
    animationStepSequences?: Array<{
        concurrent?: boolean
        nextAction?: 'none' | 'seek'
        advancesOnNext?: boolean
        rewindsOnPrevious?: boolean
    } | undefined>
    triggeredAnimations?: {
        conditions: PptxAnimationTriggerCondition[]
        restart?: 'always' | 'whenNotActive' | 'never'
        concurrent?: boolean
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

export type TextStyleAnimationProperty = 'fontWeight' | 'fontStyle' | 'fontFamily' | 'underline' | 'lineThrough' | 'textTransform' | 'outline' | 'textShadow' | 'emboss'

export type TextStyleAnimationValue = boolean | 'none' | 'normal' | 'sub' | 'super' | string

export type AnimatedVisibility = 'visible' | 'hidden'

export type AnimationOnOffProperty = 'fill.on' | 'stroke.on' | 'shadow.on'

export type AnimatedStrokeDashStyle = 'solid' | 'dot' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot'

export type AnimatedLineArrow = 'none' | 'block' | 'classic' | 'diamond' | 'oval' | 'open' | 'chevron' | 'doublechevron'

export type AnimatedLineArrowProperty = 'stroke.startArrow' | 'stroke.endArrow'

export type AnimatedLineArrowSize = 'narrow' | 'medium' | 'wide' | 'short' | 'long'

export type AnimatedLineArrowSizeProperty = 'stroke.startArrowWidth' | 'stroke.endArrowWidth' | 'stroke.startArrowLength' | 'stroke.endArrowLength'

type AnimatedImageCropProperty = 'left' | 'top' | 'right' | 'bottom'

type AnimatedImageToneProperty = 'gain' | 'blacklevel' | 'gamma'

export interface PptxAnimation {
    targetId: string
    paragraphRange?: {start: number; end: number}
    characterRange?: {start: number; end: number}
    iteration?: {intervalMs: number; ranges: Array<{start: number; end: number}>; backwards?: boolean}
    effectProgress?: number
    effect: 'appear' | 'fade' | 'wipe' | 'blinds' | 'checker' | 'randomBars' | 'strips' | 'barn' | 'dissolve' | 'shape' | 'wheel' | 'slide' | 'zoom' | 'scale' | 'motion' | 'timingOnly' | 'rotation' | 'color' | 'fontSize' | 'fontWeight' | 'strokeWidth' | 'textStyle' | 'visibility' | 'grayscale' | 'onOff' | 'strokeDash' | 'lineArrow' | 'lineArrowSize' | 'imageCrop' | 'imageTone' | 'opacity' | 'media'
    direction: 'in' | 'out'
    durationMs: number
    delayMs: number
    exclusiveGroups?: Array<{id: string; startMs: number; endMs: number}>
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
    slideFrom?: 'left' | 'right' | 'top' | 'bottom' | 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight'
    wheelSpokes?: 1 | 2 | 3 | 4 | 8
    scaleFrom?: [number, number]
    scaleTo?: [number, number]
    scaleAxis?: 'x' | 'y'
    scaleKeyframes?: Array<{offset: number; value: number}>
    scaleKeyframeMode?: 'lin' | 'discrete'
    scaleAccumulate?: boolean
    rotationFrom?: number
    rotationTo?: number
    rotationKeyframes?: Array<{offset: number; value: number}>
    rotationKeyframeMode?: 'lin' | 'discrete'
    rotationRelative?: boolean
    rotationAdditive?: boolean
    rotationAccumulate?: boolean
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
    fontSizeFormulaSamples?: Array<{offset: number; value: number}>
    fontWeightFrom?: number
    fontWeightTo?: number
    fontWeightReset?: boolean
    fontWeightKeyframes?: Array<{offset: number; value: number}>
    fontWeightKeyframeMode?: 'lin' | 'discrete'
    fontWeightFormulaSamples?: Array<{offset: number; value: number}>
    strokeWidthFrom?: number
    strokeWidthTo?: number
    strokeWidthKeyframes?: Array<{offset: number; value: number}>
    strokeWidthKeyframeMode?: 'lin' | 'discrete'
    strokeWidthFormulaSamples?: Array<{offset: number; value: number}>
    strokeDashFrom?: AnimatedStrokeDashStyle
    strokeDashTo?: AnimatedStrokeDashStyle
    strokeDashKeyframes?: Array<{offset: number; value: AnimatedStrokeDashStyle}>
    lineArrowProperty?: AnimatedLineArrowProperty
    lineArrowFrom?: AnimatedLineArrow
    lineArrowTo?: AnimatedLineArrow
    lineArrowKeyframes?: Array<{offset: number; value: AnimatedLineArrow}>
    lineArrowSizeProperty?: AnimatedLineArrowSizeProperty
    lineArrowSizeFrom?: AnimatedLineArrowSize
    lineArrowSizeTo?: AnimatedLineArrowSize
    lineArrowSizeKeyframes?: Array<{offset: number; value: AnimatedLineArrowSize}>
    textStyleProperty?: TextStyleAnimationProperty
    textStyleFrom?: TextStyleAnimationValue
    textStyleTo?: TextStyleAnimationValue
    textStyleKeyframes?: Array<{offset: number; value: TextStyleAnimationValue}>
    visibilityFrom?: AnimatedVisibility
    visibilityTo?: AnimatedVisibility
    visibilityKeyframes?: Array<{offset: number; value: AnimatedVisibility}>
    onOffProperty?: AnimationOnOffProperty
    onOffFrom?: boolean
    onOffTo?: boolean
    onOffKeyframes?: Array<{offset: number; value: boolean}>
    grayscaleFrom?: boolean
    grayscaleTo?: boolean
    grayscaleKeyframes?: Array<{offset: number; value: boolean}>
    imageCropProperty?: AnimatedImageCropProperty
    imageCropFrom?: number
    imageCropTo?: number
    imageCropKeyframes?: Array<{offset: number; value: number}>
    imageCropKeyframeMode?: 'lin' | 'discrete'
    imageToneProperty?: AnimatedImageToneProperty
    imageToneFrom?: number
    imageToneTo?: number
    imageToneKeyframes?: Array<{offset: number; value: number}>
    imageToneKeyframeMode?: 'lin' | 'discrete'
    opacityFrom?: number
    opacityTo?: number
    opacityKeyframes?: Array<{offset: number; value: number}>
    opacityKeyframeMode?: 'lin' | 'discrete'
    opacityFormulaSamples?: Array<{offset: number; value: number}>
    mediaCommand?: 'play' | 'pause' | 'togglePause' | 'stop' | 'stopAudio'
    mediaActionKey?: string
    mediaDurationMs?: number
    mediaSlideCount?: number
    mediaWaitForEnd?: boolean
    mediaWaitForEndKeys?: string[]
    mediaWaitDelayMs?: number
    mediaStartSeconds?: number
    mediaVolume?: number
    mediaMuted?: boolean
    mediaIsNarration?: boolean
    motionCoordinateProperty?: 'ppt_x' | 'ppt_y'
    motionPath?: Array<{x: number; y: number; distance: number}>
    motionPathLength?: number
    motionPathSamples?: Array<{progress: number; x: number; y: number}>
    motionAccumulate?: boolean
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

export interface PptxDocument {
    name: string
    width: number
    height: number
    slides: PptxSlide[]
    showSettings?: {loop: boolean; showAnimation: boolean; showNarration: boolean; useTimings: boolean}
    warnings: string[]
    objectUrls: string[]
    embeddedFonts: PptxEmbeddedFontFace[]
}

export interface PptxEmbeddedFontFace {
    family: string
    weight: '400' | '700'
    style: 'normal' | 'italic'
    url: string
}
