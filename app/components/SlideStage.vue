<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, useId, watch, type CSSProperties } from 'vue'
import { formatChartNumberWithColor, getChartDateAxisMajorTicks } from '~/utils/chart-number-format'
import { paragraphEntries } from '~/utils/paragraph-bullets'
import { hslToRgb, rgbToHsl } from '~/utils/color'
import { characterScript, cssFontFamily, type ScriptFont } from '~/utils/text-fonts'
import { getPptxChartStacking, isSvgPolygonGeometry, presetEllipsePoint } from '~/utils/pptx-render-utils'
import type { PptxAnimation, PptxAnimationTriggerCondition, PptxChart, PptxChartDataLabelPosition, PptxChartDataLabelStyle, PptxChartMarkerSymbol, PptxChartTickLabelPosition, PptxChartTickMark, PptxElement, PptxHyperlink, PptxParagraph, PptxShadow, PptxSlide, PptxTableCell, PptxTriggerPlayback } from '~/utils/pptx-model'
import { inverseTimingProgress, timingProgress } from '~/utils/animation-timing'
import { barnClipPath, clipPathMaskImage, dissolveClipPath, gridTransitionClipPath, gridTransitionMaskImage, randomBarClipPath, shapeEffectClipPath, stripsClipPath, wheelClipPath } from '~/utils/transitionMasks'
import type { LocaleMessageKey } from '~/composables/useLocale'

const { t } = useLocale()
const EMU_PER_POINT = 12_700
const customGradientScope = `pptx-${useId()}`.replace(/[^a-zA-Z0-9_-]/g, '_')
const svgStrokeDashPatterns: Record<string, number[]> = {
  dot: [1, 3], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3],
  dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1], lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1],
}

function svgStrokeDashPattern(dash: PptxElement['strokeDash'], cap?: PptxElement['strokeCap']): number[] | undefined {
  const pattern = Array.isArray(dash) ? dash : svgStrokeDashPatterns[dash || '']
  if (!pattern) return undefined
  const capLength = cap === 'round' || (!Array.isArray(dash) && cap === 'square') ? 1 : 0
  return pattern.map((length, index) => {
    const extension = Math.min(capLength, pattern[index % 2 ? index - 1 : index]!)
    return length + (index % 2 ? extension : -extension)
  })
}

interface CustomSvgGradientStop { offset: string; color: string; opacity: number }
interface CustomSvgLinearGradient {
  id: string
  kind: 'linear'
  x1: number
  y1: number
  x2: number
  y2: number
  stops: CustomSvgGradientStop[]
}
type CustomSvgGradientFocus =
  | { kind: 'ellipse'; cx: number; cy: number; rx: number; ry: number; color: string }
  | { kind: 'rect'; x: number; y: number; width: number; height: number; color: string }
interface CustomSvgRadialGradient {
  id: string
  kind: 'radial'
  cx: number
  cy: number
  r: number
  fx?: number
  fy?: number
  fr?: number
  colorInterpolation?: 'sRGB' | 'linearRGB'
  gradientTransform?: string
  stops: CustomSvgGradientStop[]
  focus?: CustomSvgGradientFocus
}
type CustomSvgGradient = CustomSvgLinearGradient | CustomSvgRadialGradient

interface ChartVisual {
  plot: { left: number; top: number; width: number; height: number }
  ticks: Array<{ value: number; label: string; color?: string; showLabel: boolean; x: number; y: number; labelX: number; labelY: number; labelAnchor: string }>
  xTicks: Array<{ value: number; label: string; color?: string; showLabel: boolean; x: number; y: number; labelX: number; labelY: number; labelAnchor: string }>
  axisTicks: Array<{ axis: 'category' | 'value' | 'valueSecondary' | 'xValue'; axisId?: string; minor?: boolean; x1: number; y1: number; x2: number; y2: number }>
  secondaryValueAxes: Array<{
    id: string
    horizontal: boolean
    line: { x1: number; y1: number; x2: number; y2: number }
    ticks: Array<{ value: number; label: string; color?: string; showLabel: boolean; x: number; y: number; labelX: number; labelY: number; labelAnchor: string }>
    majorGridlines: boolean
    majorTickMark?: PptxChartTickMark
    minorTickMark?: PptxChartTickMark
    minorTicks: Array<{x: number; y: number}>
    deleted?: boolean
  }>
  radarGrid: Array<{ d: string }>
  radarSpokes: Array<{ x1: number; y1: number; x2: number; y2: number }>
  categoryGridlines: Array<{ x1: number; y1: number; x2: number; y2: number }>
  categories: Array<{ text: string; color?: string; x: number; y: number; anchor: string }>
  bars: Array<{ x: number; y: number; width: number; height: number; color: string; stroke?: string; strokeWidth?: number; strokeDasharray?: string; strokeLinecap?: PptxElement['strokeCap'] }>
  areas: Array<{ d: string; color: string; stroke?: string; strokeWidth?: number; strokeDasharray?: string; strokeLinecap?: PptxElement['strokeCap']; opacity?: number }>
  lines: Array<{ d: string; color: string; width?: number; strokeDasharray?: string; strokeLinecap?: PptxElement['strokeCap'] }>
  markers: Array<{ x: number; y: number; color: string; stroke?: string; strokeWidth?: number; strokeDasharray?: string; strokeLinecap?: PptxElement['strokeCap']; radius?: number; symbol?: PptxChartMarkerSymbol }>
  slices: Array<{ d: string; color: string; stroke?: string; strokeWidth?: number; strokeDasharray?: string; strokeLinecap?: PptxElement['strokeCap'] }>
  legend: Array<ChartLegendKey & { text: string; x: number; y: number }>
  legendClip?: { id: string; x: number; y: number; width: number; height: number }
  dataLabels: Array<{ text: string; lines: string[]; x: number; y: number; anchor: string; legendKey?: ChartLegendKey; leaderLine?: { d: string; color: string }; style?: PptxChartDataLabelStyle }>
  baseline: { x1: number; y1: number; x2: number; y2: number }
  valueAxisLine?: { x1: number; y1: number; x2: number; y2: number }
}

type ChartLegendKey = {color: string; symbol?: PptxChartMarkerSymbol; line?: boolean; markerChart?: boolean}
type ChartLegendPosition = NonNullable<PptxChart['legend']>['position']

function defaultLegendOrigin(position: ChartLegendPosition | undefined, axis: 'x' | 'y') {
  if (axis === 'x') return position === 'left' ? 88 : position === 'right' ? 780 : position === 'topRight' ? 760 : 88
  return position === 'left' || position === 'right' ? 105 : position === 'top' ? 77 : position === 'topRight' ? 92 : 530
}

function chartMarkerPath(marker: ChartVisual['markers'][number]) {
  const { x, y } = marker
  const radius = marker.radius ?? 5
  const symbol = marker.symbol ?? 'circle'
  if (symbol === 'square') return `M ${x - radius} ${y - radius} h ${radius * 2} v ${radius * 2} h ${-radius * 2} Z`
  if (symbol === 'diamond') return `M ${x} ${y - radius} L ${x + radius} ${y} L ${x} ${y + radius} L ${x - radius} ${y} Z`
  if (symbol === 'triangle') return `M ${x} ${y - radius} L ${x + radius} ${y + radius} L ${x - radius} ${y + radius} Z`
  if (symbol === 'dash') return `M ${x - radius} ${y - radius * 0.22} h ${radius * 2} v ${radius * 0.44} h ${-radius * 2} Z`
  if (symbol === 'plus') {
    const arm = radius * 0.28
    return `M ${x - arm} ${y - radius} H ${x + arm} V ${y - arm} H ${x + radius} V ${y + arm} H ${x + arm} V ${y + radius} H ${x - arm} V ${y + arm} H ${x - radius} V ${y - arm} H ${x - arm} Z`
  }
  if (symbol === 'x') {
    const arm = radius * 0.28
    return `M ${x - radius} ${y - radius + arm} L ${x - radius + arm} ${y - radius} L ${x} ${y - arm} L ${x + radius - arm} ${y - radius} L ${x + radius} ${y - radius + arm} L ${x + arm} ${y} L ${x + radius} ${y + radius - arm} L ${x + radius - arm} ${y + radius} L ${x} ${y + arm} L ${x - radius + arm} ${y + radius} L ${x - radius} ${y + radius - arm} L ${x - arm} ${y} Z`
  }
  if (symbol === 'star') {
    return Array.from({ length: 10 }, (_, index) => {
      const angle = -Math.PI / 2 + index * Math.PI / 5
      const pointRadius = radius * (index % 2 ? 0.45 : 1)
      return `${index ? 'L' : 'M'} ${x + Math.cos(angle) * pointRadius} ${y + Math.sin(angle) * pointRadius}`
    }).join(' ') + ' Z'
  }
  return ''
}

// shortcut: uniform Catmull-Rom parameter spacing; upgrade if PowerPoint reference decks show a different curve.
function smoothChartLinePath(points: Array<{x: number; y: number}>, smooth: boolean) {
  const first = points[0]
  if (!first) return ''

  let path = `M ${first.x} ${first.y}`
  for (let index = 0; index < points.length - 1; index++) {
    const start = points[index]
    const end = points[index + 1]
    if (!start || !end) continue
    if (!smooth || points.length < 3) {
      path += ` L ${end.x} ${end.y}`
      continue
    }

    const previous = points[Math.max(0, index - 1)] ?? start
    const next = points[Math.min(index + 2, points.length - 1)] ?? end
    path += ` C ${start.x + (end.x - previous.x) / 6} ${start.y + (end.y - previous.y) / 6} ${end.x - (next.x - start.x) / 6} ${end.y - (next.y - start.y) / 6} ${end.x} ${end.y}`
  }
  return path
}

function makeChartVisual(chart: PptxChart, elementId: string, chartFrameWidthEmu: number): ChartVisual {
  const chartStrokeWidth = (widthEmu?: number) => widthEmu === undefined ? undefined : ((widthEmu || 9_525) * 1000 / Math.max(chartFrameWidthEmu, 1))
  const chartStrokeStyle = (series: PptxChart['series'][number], index: number | undefined, width: number) => {
    const dash = index === undefined ? series.strokeDash : series.pointStrokeDashes?.[index] ?? series.strokeDash
    const cap = index === undefined ? series.strokeCap : series.pointStrokeCaps?.[index] ?? series.strokeCap
    const pattern = svgStrokeDashPattern(dash, cap)
    return {
      ...(pattern ? {strokeDasharray: pattern.map(length => length * width).join(' ')} : {}),
      ...(cap ? {strokeLinecap: cap} : {}),
    }
  }
  const combo = chart.series.some(series => series.chartType)
  const seriesType = (series: PptxChart['series'][number]) => series.chartType || chart.type
  const axisEntryForSeries = (series: PptxChart['series'][number]) => chart.valueAxes?.find(item => item.id === series.valueAxisId)
  const primaryAxisId = chart.valueAxes?.find(item => item.primary)?.id
  const axisValuesForId = (axisId?: string) => chart.series
    .filter(series => !axisId || (series.valueAxisId || primaryAxisId) === axisId)
    .flatMap(series => series.values.filter((value): value is number => value !== null && value !== undefined))
  const scatter = chart.type === 'scatter'
  const bubble = chart.type === 'bubble'
  const radar = chart.type === 'radar'
  const stock = chart.type === 'stock' ? chart.stock : undefined
  const xyChart = scatter || bubble
  const position = chart.legend?.position
  const plotPosition = chart.legend?.overlay ? undefined : position
  const plot = plotPosition === 'left' ? { left: 330, top: 75, width: 608, height: 390 }
    : plotPosition === 'right' ? { left: 88, top: 75, width: 650, height: 390 }
      : plotPosition === 'top' ? { left: 88, top: 145, width: 850, height: 320 }
        : { left: 88, top: 75, width: 850, height: 390 }
  const radarCenterX = plot.left + plot.width / 2
  const radarCenterY = plot.top + plot.height / 2
  const radarRadiusX = plot.width * 0.38
  const radarRadiusY = plot.height * 0.38
  const radarPoint = (index: number, fraction = 1, padding = 0) => {
    const angle = -Math.PI / 2 + Math.PI * 2 * index / Math.max(chart.categories.length, 1)
    return {
      x: radarCenterX + Math.cos(angle) * (radarRadiusX * fraction + padding),
      y: radarCenterY + Math.sin(angle) * (radarRadiusY * fraction + padding),
      angle,
    }
  }
  const values = stock
    ? [stock.highIndex, stock.lowIndex, stock.closeIndex, stock.openIndex].filter((index): index is number => index !== undefined).flatMap(index => chart.series[index]?.values.filter((value): value is number => value !== null && value !== undefined) || [])
    : combo ? axisValuesForId(primaryAxisId)
      : chart.series.flatMap(series => series.values.filter((value): value is number => value !== null && value !== undefined))
  const stacking = !combo && ['bar', 'line', 'area'].includes(chart.type) && chart.grouping && chart.grouping !== 'clustered'
    ? getPptxChartStacking(chart.series, chart.grouping)
    : undefined
  const stackedSegments = new Map<number, ReturnType<typeof getPptxChartStacking>['segments'][number]>()
  if (combo) {
    const groupIndexes = new Map<number, number[]>()
    chart.series.forEach((series, index) => {
      const groupIndex = series.chartGroupIndex ?? 0
      groupIndexes.set(groupIndex, [...(groupIndexes.get(groupIndex) || []), index])
    })
    for (const [groupIndex, indexes] of groupIndexes) {
      const grouping = chart.series[indexes[0]!]!.grouping || 'clustered'
      if (grouping === 'clustered') continue
      const groupSeries = indexes.map(index => chart.series[index]!)
      const groupStacking = getPptxChartStacking(groupSeries, grouping)
      indexes.forEach((index, seriesIndex) => stackedSegments.set(index, groupStacking.segments[seriesIndex]!))
    }
  }
  const stackedSegmentFor = (seriesIndex: number) => combo ? stackedSegments.get(seriesIndex) : stacking?.segments[seriesIndex]
  const axisScaleFraction = (value: number, minimum: number, maximum: number, axis: PptxChart['valueAxis']) => {
    if (!axis?.logBase) return (value - minimum) / (maximum - minimum)
    const positiveValue = value > 0 ? value : minimum
    return (Math.log(positiveValue) - Math.log(minimum)) / (Math.log(maximum) - Math.log(minimum))
  }
  const axisMajorTickValues = (axis: PptxChart['valueAxis'], minimum: number, maximum: number) => {
    if (!axis?.logBase) {
      const unit = axis?.majorUnit
      return unit
        ? Array.from({length: Math.min(101, Math.floor((maximum - minimum) / unit + 1e-9) + 1)}, (_, index) => Number((minimum + index * unit).toPrecision(12)))
        : Array.from({length: 5}, (_, index) => minimum + (maximum - minimum) * index / 4)
    }
    const logarithmBase = Math.log(axis.logBase)
    const minimumExponent = Math.log(minimum) / logarithmBase
    const maximumExponent = Math.log(maximum) / logarithmBase
    const step = axis.majorUnit ?? 1
    const firstExponent = Math.ceil((minimumExponent - 1e-9) / step) * step
    const tickCount = Math.min(101, Math.floor((maximumExponent - firstExponent) / step + 1e-9) + 1)
    const values = Array.from({length: Math.max(0, tickCount)}, (_, index) => axis.logBase! ** (firstExponent + index * step))
      .filter(value => Number.isFinite(value) && value >= minimum * (1 - 1e-9) && value <= maximum * (1 + 1e-9))
    return values.length ? values : [minimum, maximum]
  }
  const axisBoundsForSeries = (series: PptxChart['series'][number]) => {
    const axis = axisEntryForSeries(series)?.axis || chart.valueAxis
    if (!combo && stacking) {
      const minimum = chart.valueAxis?.min ?? stacking.min
      const upper = chart.valueAxis?.max ?? (stacking.max === minimum ? minimum + 1 : stacking.max)
      return {axis, minimum, upper}
    }
    const values = axisValuesForId(series.valueAxisId || primaryAxisId)
    const dataMin = values.length ? axis?.logBase ? Math.min(...values) : Math.min(0, ...values) : axis?.logBase ? 1 : 0
    const dataMax = values.length ? Math.max(...values, dataMin) : axis?.logBase ? axis.logBase : 1
    const minimum = axis?.min ?? dataMin
    const fallbackUpper = axis?.logBase ? minimum * axis.logBase : minimum + 1
    return {axis, minimum, upper: axis?.max ?? (dataMax === minimum ? fallbackUpper : dataMax)}
  }
  const seriesValuePosition = (series: PptxChart['series'][number], value: number) => {
    const {axis, minimum: axisMinimum, upper: axisUpper} = axisBoundsForSeries(series)
    const fraction = axisScaleFraction(value, axisMinimum, axisUpper, axis)
    const axisFraction = axis?.reverse ? 1 - fraction : fraction
    return chart.direction === 'horizontal'
      ? plot.left + plot.width * axisFraction
      : plot.top + plot.height * (1 - axisFraction)
  }
  const dataMin = stacking?.min ?? (xyChart && values.length ? Math.min(...values) : Math.min(0, ...values))
  const dataMax = stacking?.max ?? (xyChart && values.length ? Math.max(...values) : Math.max(0, ...values))
  const minimum = chart.valueAxis?.min ?? dataMin
  const upper = chart.valueAxis?.max ?? (dataMax === dataMin ? dataMin + 1 : dataMax)
  const xValues = chart.series.flatMap(series => (series.xValues || []).filter((value): value is number => value !== null && value !== undefined))
  const xDataMin = xyChart && xValues.length ? Math.min(...xValues) : 0
  const xDataMax = xyChart && xValues.length ? Math.max(...xValues) : 1
  const xMinimum = chart.xAxis?.min ?? (xDataMin === xDataMax ? xDataMin - (Math.abs(xDataMin) * 0.05 || 0.5) : xDataMin)
  const xUpper = chart.xAxis?.max ?? (xDataMin === xDataMax ? xDataMax + (Math.abs(xDataMax) * 0.05 || 0.5) : xDataMax)
  const categoryCount = Math.max(chart.categories.length, 1)
  const formatCategoryAxisLabel = (text: string) => {
    if (!chart.categoryAxisNumberFormat || !text.trim()) return {text}
    const value = Number(text)
    return Number.isFinite(value) ? formatChartNumberWithColor(value, chart.categoryAxisNumberFormat, chart.date1904, chart.indexedColors) : {text}
  }
  const horizontalValueAxis = chart.direction === 'horizontal'
  const valueAxisSide = chart.valueAxis?.position ?? (horizontalValueAxis ? 'b' : 'l')
  const xAxisSide = chart.xAxis?.position ?? 'b'
  const categoryAxisSide = chart.categoryAxisPosition ?? (horizontalValueAxis ? 'l' : 'b')
  const valueTickLabelPosition = chart.valueAxis?.tickLabelPosition ?? 'nextTo'
  const categoryTickLabelPosition = chart.categoryAxisTickLabelPosition ?? 'nextTo'
  const categoryAxisLabelStep = chart.categoryAxisTickLabelSkip ?? 1
  const categoryAxisTickStep = chart.categoryAxisTickMarkSkip ?? 1
  const showCategoryAxisLabel = (index: number) => index % categoryAxisLabelStep === 0
  const showCategoryAxisTick = (index: number) => index % categoryAxisTickStep === 0
  const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value))
  const dateValues = chart.dateAxis?.values
  const dateMinimum = chart.dateAxis?.min ?? (dateValues?.length ? Math.min(...dateValues) : 0)
  const dateMaximum = chart.dateAxis?.max ?? (dateValues?.length ? Math.max(...dateValues) : 1)
  const hasDateAxisScale = dateValues?.length === categoryCount && dateMaximum > dateMinimum
  const dateAxisMajorTickValues = hasDateAxisScale && chart.dateAxis?.majorUnit !== undefined && chart.dateAxis.majorTimeUnit
    ? getChartDateAxisMajorTicks(dateMinimum, dateMaximum, chart.dateAxis.majorUnit, chart.dateAxis.majorTimeUnit, chart.date1904)
    : undefined
  const dateAxisMinorTickValues = hasDateAxisScale && chart.dateAxis?.minorUnit !== undefined && chart.dateAxis.minorTimeUnit
    ? getChartDateAxisMajorTicks(dateMinimum, dateMaximum, chart.dateAxis.minorUnit, chart.dateAxis.minorTimeUnit, chart.date1904)
    : undefined
  const dateAxisFraction = (value: number) => {
    const fraction = clamp((value - dateMinimum) / (dateMaximum - dateMinimum), 0, 1)
    return chart.dateAxis?.reverse ? 1 - fraction : fraction
  }
  const categoryFraction = (index: number) => {
    if (!hasDateAxisScale || dateValues?.[index] === undefined) return (index + 0.5) / categoryCount
    return dateAxisFraction(dateValues[index]!)
  }
  const categoryX = (index: number) => plot.left + plot.width * categoryFraction(index)
  const categoryY = (index: number) => plot.top + plot.height * categoryFraction(index)
  const categoryBandSize = (index: number) => {
    const dimension = horizontalValueAxis ? plot.height : plot.width
    if (!hasDateAxisScale) return dimension / categoryCount
    const current = categoryFraction(index)
    const gaps = [
      index > 0 ? Math.abs(current - categoryFraction(index - 1)) : undefined,
      index + 1 < categoryCount ? Math.abs(categoryFraction(index + 1) - current) : undefined,
    ].filter((gap): gap is number => gap !== undefined && gap > 0)
    return Math.max(1, dimension * (gaps.length ? Math.min(...gaps) : 1 / categoryCount))
  }
  const axisFraction = (value: number) => {
    const fraction = axisScaleFraction(value, minimum, upper, chart.valueAxis)
    return chart.valueAxis?.reverse ? 1 - fraction : fraction
  }
  const xAxisFraction = (value: number) => {
    const fraction = axisScaleFraction(value, xMinimum, xUpper, chart.xAxis)
    return chart.xAxis?.reverse ? 1 - fraction : fraction
  }
  const xPosition = (value: number) => plot.left + plot.width * xAxisFraction(value)
  const valuePosition = (value: number) => plot.top + plot.height * (1 - axisFraction(value))
  const withoutMajorTickValues = (values: number[], minimum: number, maximum: number, majorValues: number[]) => values.filter(value =>
    value > minimum + Number.EPSILON * Math.abs(minimum) && value < maximum - Number.EPSILON * Math.abs(maximum)
      && !majorValues.some(major => Math.abs(value - major) <= 1e-9 * Math.max(Math.abs(value), Math.abs(major), 1e-300)),
  )
  const minorAxisTickValues = (majorValues: number[], minimum: number, maximum: number, minorUnit?: number, logBase?: number) => {
    const values: number[] = []
    if (logBase) {
      const logarithmBase = Math.log(logBase)
      const minimumExponent = Math.log(minimum) / logarithmBase
      const maximumExponent = Math.log(maximum) / logarithmBase
      if (minorUnit && Number.isFinite(minorUnit) && minorUnit > 0) {
        const firstExponent = (Math.floor(minimumExponent / minorUnit + 1e-9) + 1) * minorUnit
        const count = Math.min(512, Math.floor((maximumExponent - firstExponent) / minorUnit + 1e-9) + 1)
        for (let index = 0; index < count; index++) values.push(logBase ** (firstExponent + index * minorUnit))
      } else if (Number.isInteger(logBase) && logBase <= 16) {
        for (let exponent = Math.floor(minimumExponent); exponent < maximumExponent && values.length < 512; exponent++) {
          const decade = logBase ** exponent
          for (let multiplier = 2; multiplier < logBase && values.length < 512; multiplier++) values.push(multiplier * decade)
        }
      } else {
        for (let index = 1; index < majorValues.length && values.length < 512; index++) {
          const start = majorValues[index - 1]!
          const end = majorValues[index]!
          for (let part = 1; part < 5 && values.length < 512; part++) values.push(Math.exp(Math.log(start) + (Math.log(end) - Math.log(start)) * part / 5))
        }
      }
    } else if (minorUnit && Number.isFinite(minorUnit) && minorUnit > 0) {
      const count = Math.min(512, Math.floor((maximum - minimum) / minorUnit + 1e-9))
      for (let index = 1; index <= count; index++) values.push(Number((minimum + index * minorUnit).toPrecision(12)))
    } else {
      for (let index = 1; index < majorValues.length; index++) {
        const start = majorValues[index - 1]!
        const end = majorValues[index]!
        for (let part = 1; part < 5; part++) values.push(start + (end - start) * part / 5)
      }
    }
    return withoutMajorTickValues(values, minimum, maximum, majorValues)
  }
  const categoryCrossingFraction = chart.categoryAxisCrossesAt !== undefined
    ? (clamp(chart.categoryAxisCrossesAt, 1, categoryCount) - 0.5) / categoryCount
    : chart.valueAxis?.crosses === 'min' ? 0 : chart.valueAxis?.crosses === 'max' ? 1 : undefined
  const xyValueAxisCrossingX = xPosition(clamp(chart.valueAxis?.crossesAt ?? (chart.valueAxis?.crosses === 'min' ? xMinimum : chart.valueAxis?.crosses === 'max' ? xUpper : clamp(0, xMinimum, xUpper)), xMinimum, xUpper))
  const xyXAxisCrossingY = valuePosition(clamp(chart.xAxis?.crossesAt ?? (chart.xAxis?.crosses === 'min' ? minimum : chart.xAxis?.crosses === 'max' ? upper : clamp(0, minimum, upper)), minimum, upper))
  const valueAxisCrossingPosition = categoryCrossingFraction === undefined ? undefined
    : horizontalValueAxis ? plot.top + plot.height * categoryCrossingFraction
      : plot.left + plot.width * categoryCrossingFraction
  const categoryAxisCrossingValue = chart.valueAxis?.crossesAt ?? (chart.categoryAxisCrosses === 'min' ? minimum : chart.categoryAxisCrosses === 'max' ? upper : clamp(0, minimum, upper))
  const categoryAxisCrossingCoordinate = horizontalValueAxis
    ? plot.left + plot.width * axisFraction(clamp(categoryAxisCrossingValue, minimum, upper))
    : valuePosition(clamp(categoryAxisCrossingValue, minimum, upper))
  const valueAxisLineCoordinate = xyChart ? xyValueAxisCrossingX : valueAxisCrossingPosition ?? (horizontalValueAxis
    ? valueAxisSide === 't' ? plot.top : plot.top + plot.height
    : valueAxisSide === 'r' ? plot.left + plot.width : plot.left)
  const valueAxisLine = xyChart
    ? { x1: valueAxisLineCoordinate, y1: plot.top, x2: valueAxisLineCoordinate, y2: plot.top + plot.height }
    : horizontalValueAxis
      ? { x1: plot.left, y1: valueAxisLineCoordinate, x2: plot.left + plot.width, y2: valueAxisLineCoordinate }
      : { x1: valueAxisLineCoordinate, y1: plot.top, x2: valueAxisLineCoordinate, y2: plot.top + plot.height }
  const tickValues = axisMajorTickValues(chart.valueAxis, minimum, upper)
  const ticks = tickValues.map(value => {
    const x = radar ? radarCenterX : plot.left + plot.width * axisFraction(value)
    const y = radar ? radarCenterY - radarRadiusY * axisFraction(value) : valuePosition(value)
    const formattedLabel = formatChartNumberWithColor(value, chart.valueAxis?.numberFormat ?? (chart.grouping === 'percentStacked' ? '0%' : undefined), chart.date1904, chart.indexedColors)
    return {
      value,
      label: formattedLabel.text,
      ...(formattedLabel.color ? {color: formattedLabel.color} : {}),
      showLabel: valueTickLabelPosition !== 'none' && !chart.valueAxis?.deleted,
      x,
      y,
      labelX: radar ? radarCenterX - 10 : horizontalValueAxis ? x : valueTickLabelPosition === 'high' ? plot.left + plot.width + 10 : valueTickLabelPosition === 'low' ? plot.left - 10 : valueAxisLineCoordinate + (valueAxisSide === 'r' ? 10 : -10),
      labelY: radar ? y + 5 : horizontalValueAxis ? valueTickLabelPosition === 'high' ? plot.top + plot.height + 22 : valueTickLabelPosition === 'low' ? plot.top - 10 : valueAxisLineCoordinate + (valueAxisSide === 't' ? -10 : 22) : y + 5,
      labelAnchor: radar ? 'end' : horizontalValueAxis ? 'middle' : valueTickLabelPosition === 'high' ? 'start' : valueTickLabelPosition === 'low' ? 'end' : valueAxisSide === 'r' ? 'start' : 'end',
    }
  })
  const secondaryValueAxes: ChartVisual['secondaryValueAxes'] = (chart.valueAxes || []).filter(item => !item.primary).map(({id, axis}) => {
    const values = axisValuesForId(id)
    const minimum = axis.min ?? (values.length ? Math.min(0, ...values) : 0)
    const upper = axis.max ?? Math.max(minimum + 1, 0, ...values)
    const horizontal = horizontalValueAxis
    const side = axis.position ?? (horizontal ? 't' : 'r')
    const lineCoordinate = horizontal
      ? side === 't' ? plot.top : plot.top + plot.height
      : side === 'l' ? plot.left : plot.left + plot.width
    const axisLine = horizontal
      ? {x1: plot.left, y1: lineCoordinate, x2: plot.left + plot.width, y2: lineCoordinate}
      : {x1: lineCoordinate, y1: plot.top, x2: lineCoordinate, y2: plot.top + plot.height}
    const axisTickValues = axisMajorTickValues(axis, minimum, upper)
    const tickPosition = (value: number) => {
      const fraction = axisScaleFraction(value, minimum, upper, axis)
      const normalized = axis.reverse ? 1 - fraction : fraction
      return horizontal ? plot.left + plot.width * normalized : plot.top + plot.height * (1 - normalized)
    }
    const labelPosition = axis.tickLabelPosition ?? 'nextTo'
    const axisTicks = axisTickValues.map(value => {
      const position = tickPosition(value)
      const formattedLabel = formatChartNumberWithColor(value, axis.numberFormat, chart.date1904, chart.indexedColors)
      return {
        value,
        label: formattedLabel.text,
        ...(formattedLabel.color ? {color: formattedLabel.color} : {}),
        showLabel: !axis.deleted && labelPosition !== 'none',
        x: horizontal ? position : lineCoordinate,
        y: horizontal ? lineCoordinate : position,
        labelX: horizontal ? position : lineCoordinate + (side === 'l' ? -10 : 10),
        labelY: horizontal ? lineCoordinate + (side === 't' ? -10 : 22) : position + 5,
        labelAnchor: horizontal ? 'middle' : side === 'l' ? 'end' : 'start',
      }
    })
    const minorTicks = axis.minorTickMark && axis.minorTickMark !== 'none'
      ? minorAxisTickValues(axisTickValues, minimum, upper, axis.minorUnit, axis.logBase).map(value => {
          const position = tickPosition(value)
          return {x: horizontal ? position : lineCoordinate, y: horizontal ? lineCoordinate : position}
        })
      : []
    return {id, horizontal, line: axisLine, ticks: axisTicks, minorTicks, majorGridlines: !!axis.majorGridlines, ...(axis.majorTickMark ? {majorTickMark: axis.majorTickMark} : {}), ...(axis.minorTickMark ? {minorTickMark: axis.minorTickMark} : {}), ...(axis.deleted ? {deleted: true} : {})}
  })
  const xTickValues = axisMajorTickValues(chart.xAxis, xMinimum, xUpper)
  const xTickLabelPosition = chart.xAxis?.tickLabelPosition ?? 'nextTo'
  const xTicks = xyChart ? xTickValues.map(value => {
    const x = xPosition(value)
    const formattedLabel = formatChartNumberWithColor(value, chart.xAxis?.numberFormat, chart.date1904, chart.indexedColors)
    return {
      value,
      label: formattedLabel.text,
      ...(formattedLabel.color ? {color: formattedLabel.color} : {}),
      showLabel: xTickLabelPosition !== 'none' && !chart.xAxis?.deleted,
      x,
      y: xyXAxisCrossingY,
      labelX: x,
      labelY: xTickLabelPosition === 'high' ? plot.top - 10
        : xTickLabelPosition === 'low' ? plot.top + plot.height + 22
          : xyXAxisCrossingY + (xAxisSide === 't' ? -10 : 22),
      labelAnchor: 'middle',
    }
  }) : []
  const axisTicks: ChartVisual['axisTicks'] = []
  const radarGrid: ChartVisual['radarGrid'] = []
  const radarSpokes: ChartVisual['radarSpokes'] = []
  const categoryGridlines: ChartVisual['categoryGridlines'] = []
  const addAxisTick = (axis: 'category' | 'value' | 'valueSecondary' | 'xValue', mark: PptxChartTickMark | undefined, x: number, y: number, horizontal: boolean, inwardDirection: number, axisId?: string, minor = false) => {
    if (!mark || mark === 'none') return
    const length = minor ? 3 : 5
    const start = mark === 'cross' ? -length : 0
    const end = mark === 'cross' ? length : mark === 'in' ? length : -length
    axisTicks.push({
      axis,
      ...(axisId ? {axisId} : {}),
      ...(minor ? {minor: true} : {}),
      x1: horizontal ? x + start * inwardDirection : x,
      y1: horizontal ? y : y + start * inwardDirection,
      x2: horizontal ? x + end * inwardDirection : x,
      y2: horizontal ? y : y + end * inwardDirection,
    })
  }
  const valueTickInwardDirection = horizontalValueAxis ? valueAxisSide === 't' ? 1 : -1 : valueAxisSide === 'l' ? 1 : -1
  if (!chart.valueAxis?.deleted) {
    if (radar) for (const tick of ticks) addAxisTick('value', chart.valueAxis?.majorTickMark, radarCenterX, tick.y, true, 1)
    else for (const tick of ticks) addAxisTick('value', chart.valueAxis?.majorTickMark, horizontalValueAxis ? tick.x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : tick.y, !horizontalValueAxis, valueTickInwardDirection)
    if (chart.valueAxis?.minorTickMark && chart.valueAxis.minorTickMark !== 'none') {
      for (const value of minorAxisTickValues(tickValues, minimum, upper, chart.valueAxis.minorUnit, chart.valueAxis.logBase)) {
        const x = radar ? radarCenterX : plot.left + plot.width * axisFraction(value)
        const y = radar ? radarCenterY - radarRadiusY * axisFraction(value) : valuePosition(value)
        if (radar) addAxisTick('value', chart.valueAxis.minorTickMark, radarCenterX, y, true, 1, undefined, true)
        else addAxisTick('value', chart.valueAxis.minorTickMark, horizontalValueAxis ? x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : y, !horizontalValueAxis, valueTickInwardDirection, undefined, true)
      }
    }
  }
  if (xyChart && !chart.xAxis?.deleted) {
    const inwardDirection = xAxisSide === 't' ? 1 : -1
    for (const tick of xTicks) addAxisTick('xValue', chart.xAxis?.majorTickMark, tick.x, tick.y, true, inwardDirection)
    if (chart.xAxis?.minorTickMark && chart.xAxis.minorTickMark !== 'none') {
      for (const value of minorAxisTickValues(xTickValues, xMinimum, xUpper, chart.xAxis.minorUnit, chart.xAxis.logBase)) addAxisTick('xValue', chart.xAxis.minorTickMark, xPosition(value), xyXAxisCrossingY, true, inwardDirection, undefined, true)
    }
  }
  for (const axis of secondaryValueAxes) {
    if (axis.deleted) continue
    const inwardDirection = axis.horizontal
      ? axis.line.y1 === plot.top ? 1 : -1
      : axis.line.x1 === plot.left ? 1 : -1
    for (const tick of axis.ticks) addAxisTick('valueSecondary', axis.majorTickMark, tick.x, tick.y, !axis.horizontal, inwardDirection, axis.id)
    for (const tick of axis.minorTicks) addAxisTick('valueSecondary', axis.minorTickMark, tick.x, tick.y, !axis.horizontal, inwardDirection, axis.id, true)
  }
  const categoryTickInwardDirection = horizontalValueAxis ? categoryAxisSide === 'l' ? 1 : -1 : categoryAxisSide === 't' ? 1 : -1
  if (!xyChart && !chart.categoryAxisDeleted && chart.categoryAxisMajorTickMark && chart.categoryAxisMajorTickMark !== 'none') {
    if (radar) chart.categories.forEach((_category, index) => {
      if (!showCategoryAxisTick(index)) return
      const point = radarPoint(index)
      const start = chart.categoryAxisMajorTickMark === 'cross' ? -5 : chart.categoryAxisMajorTickMark === 'in' ? -5 : 0
      const end = chart.categoryAxisMajorTickMark === 'cross' ? 5 : chart.categoryAxisMajorTickMark === 'in' ? 0 : 5
      axisTicks.push({axis: 'category', x1: point.x + Math.cos(point.angle) * start, y1: point.y + Math.sin(point.angle) * start, x2: point.x + Math.cos(point.angle) * end, y2: point.y + Math.sin(point.angle) * end})
    })
    else for (let index = 0; index < (dateAxisMajorTickValues?.length ?? categoryCount); index++) {
      if (!showCategoryAxisTick(index)) continue
      const fraction = dateAxisMajorTickValues ? dateAxisFraction(dateAxisMajorTickValues[index]!) : undefined
      const centerX = fraction === undefined ? categoryX(index) : plot.left + plot.width * fraction
      const centerY = fraction === undefined ? categoryY(index) : plot.top + plot.height * fraction
      addAxisTick('category', chart.categoryAxisMajorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : centerX, horizontalValueAxis ? centerY : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection)
    }
  }
  if (!xyChart && !chart.categoryAxisDeleted && chart.categoryAxisMinorTickMark && chart.categoryAxisMinorTickMark !== 'none') {
    if (radar) chart.categories.forEach((_category, index) => {
      const current = radarPoint(index)
      const angle = current.angle + Math.PI / chart.categories.length
      const x = radarCenterX + Math.cos(angle) * radarRadiusX
      const y = radarCenterY + Math.sin(angle) * radarRadiusY
      const start = chart.categoryAxisMinorTickMark === 'cross' ? -3 : chart.categoryAxisMinorTickMark === 'in' ? -3 : 0
      const end = chart.categoryAxisMinorTickMark === 'cross' ? 3 : chart.categoryAxisMinorTickMark === 'in' ? 0 : 3
      axisTicks.push({axis: 'category', minor: true, x1: x + Math.cos(angle) * start, y1: y + Math.sin(angle) * start, x2: x + Math.cos(angle) * end, y2: y + Math.sin(angle) * end})
    })
    else if (hasDateAxisScale) {
      const values = dateAxisMinorTickValues
        ? withoutMajorTickValues(dateAxisMinorTickValues, dateMinimum, dateMaximum, dateAxisMajorTickValues || [])
        : minorAxisTickValues(dateAxisMajorTickValues || dateValues || [], dateMinimum, dateMaximum)
      for (const value of values) {
        const fraction = dateAxisFraction(value)
        addAxisTick('category', chart.categoryAxisMinorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : plot.left + plot.width * fraction, horizontalValueAxis ? plot.top + plot.height * fraction : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection, undefined, true)
      }
    } else {
      for (let index = 0; index < categoryCount - 1; index++) {
        const centerX = (categoryX(index) + categoryX(index + 1)) / 2
        const centerY = (categoryY(index) + categoryY(index + 1)) / 2
        addAxisTick('category', chart.categoryAxisMinorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : centerX, horizontalValueAxis ? centerY : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection, undefined, true)
      }
    }
  }
  if (!xyChart && !radar && !chart.categoryAxisDeleted && chart.categoryAxisMajorGridlines) {
    const positions = dateAxisMajorTickValues
      ? dateAxisMajorTickValues.map(value => dateAxisFraction(value))
      : Array.from({ length: categoryCount }, (_, index) => categoryFraction(index))
    for (const fraction of positions) {
      if (horizontalValueAxis) {
        const y = plot.top + plot.height * fraction
        categoryGridlines.push({ x1: plot.left, y1: y, x2: plot.left + plot.width, y2: y })
      } else {
        const x = plot.left + plot.width * fraction
        categoryGridlines.push({ x1: x, y1: plot.top, x2: x, y2: plot.top + plot.height })
      }
    }
  }
  const categories: ChartVisual['categories'] = []
  const bars: ChartVisual['bars'] = []
  const areas: ChartVisual['areas'] = []
  const lines: ChartVisual['lines'] = []
  const markers: ChartVisual['markers'] = []
  const slices: ChartVisual['slices'] = []
  const legend: ChartVisual['legend'] = []
  const dataLabels: ChartVisual['dataLabels'] = []
  if (combo) chart.series.forEach((series, seriesIndex) => {
    if (seriesType(series) !== 'area') return
    const axis = axisEntryForSeries(series)?.axis || chart.valueAxis
    const {minimum, upper} = axisBoundsForSeries(series)
    const lowerValue = axis?.crossesAt ?? (axis?.crosses === 'max' ? upper : axis?.crosses === 'min' ? minimum : clamp(0, minimum, upper))
    let upperPoints: Array<{x: number; y: number}> = []
    let lowerPoints: Array<{x: number; y: number}> = []
    const flushArea = () => {
      if (upperPoints.length > 1) {
        const points = [...upperPoints, ...lowerPoints.reverse()]
        areas.push({d: `${points.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')} Z`, color: series.color, stroke: series.stroke ?? series.color, strokeWidth: chartStrokeWidth(series.strokeWidthEmu), ...chartStrokeStyle(series, undefined, chartStrokeWidth(series.strokeWidthEmu) ?? 1)})
      }
      upperPoints = []
      lowerPoints = []
    }
    for (let index = 0; index < categoryCount; index++) {
      const value = series.values[index]
      if (value == null) {
        flushArea()
        continue
      }
      const stackedSegment = stackedSegmentFor(seriesIndex)?.[index]
      const centerX = categoryX(index)
      const centerY = categoryY(index)
      upperPoints.push(horizontalValueAxis
        ? {x: seriesValuePosition(series, stackedSegment?.end ?? value), y: centerY}
        : {x: centerX, y: seriesValuePosition(series, stackedSegment?.end ?? value)})
      lowerPoints.push(horizontalValueAxis
        ? {x: seriesValuePosition(series, stackedSegment?.start ?? lowerValue), y: centerY}
        : {x: centerX, y: seriesValuePosition(series, stackedSegment?.start ?? lowerValue)})
    }
    flushArea()
  })
  const formatValue = (value: number, series: PptxChart['series'][number], index: number) => formatChartNumberWithColor(value, series.valueNumberFormatOverrides[index] ?? series.numberFormat ?? chart.numberFormat, chart.date1904, chart.indexedColors)
  const labelPosition = (series: PptxChart['series'][number], index: number, fallback: PptxChartDataLabelPosition) => series.dataLabelPositionOverrides[index] ?? series.dataLabelPosition ?? chart.dataLabelPosition ?? fallback
  const chartLegendKey = (series: PptxChart['series'][number], index?: number): ChartLegendKey => {
    const color = index === undefined ? series.color : series.pointColors[index] || series.color
    const type = seriesType(series)
    if (!['line', 'scatter', 'radar'].includes(type)) return {color}
    const symbol = index === undefined ? series.markerSymbol : series.markerSymbolOverrides?.[index] ?? series.markerSymbol
    const showMarker = symbol === undefined
      ? type === 'line' ? series.lineShowMarker !== false : type === 'scatter' ? series.scatterShowMarker === true : series.radarShowMarker === true
      : symbol !== 'none'
    const showLine = type === 'scatter' ? series.scatterShowLine === true : type === 'line' ? series.lineShowLine !== false : series.radarShowLine !== false
    return {
      color,
      markerChart: true,
      ...(showMarker ? {symbol: symbol ?? 'circle'} : {}),
      ...(showLine ? {line: true} : {}),
    }
  }
  const pushDataLabel = (label: ReturnType<typeof formatChartNumberWithColor>, x: number, y: number, anchor: string, series: PptxChart['series'][number], index: number, leaderLine?: string) => {
    const style = {...series.dataLabelStyle, ...series.dataLabelStyleOverrides[index]}
    if (label.color) style.color = label.color
    const lines = label.text.split(/\r?\n/)
    const showLegendKey = series.legendKeyOverrides[index] ?? series.showLegendKey
    dataLabels.push({ text: label.text, lines, x, y, anchor, ...(showLegendKey ? {legendKey: chartLegendKey(series, index)} : {}), ...(leaderLine ? {leaderLine: {d: leaderLine, color: style.color || '#64748b'}} : {}), ...(Object.keys(style).length ? {style} : {}) })
  }
  const formatDataLabel = (series: PptxChart['series'][number], index: number, value: number, category: string, total?: number, bubbleSize?: number) => {
    const showValue = series.valueLabelOverrides[index] ?? series.showValueLabels
    const showLegendKey = series.legendKeyOverrides[index] ?? series.showLegendKey
    const showCategory = series.categoryNameLabelOverrides[index] ?? series.showCategoryNameLabels
    const showSeries = series.seriesNameLabelOverrides[index] ?? series.showSeriesNameLabels
    const isPieLike = chart.type === 'pie' || chart.type === 'doughnut'
    const showPercent = isPieLike && (series.percentLabelOverrides[index] ?? series.showPercentLabels)
    const showBubbleSize = bubble && bubbleSize !== undefined && (series.bubbleSizeLabelOverrides[index] ?? series.showBubbleSizeLabels)
    const parts = [
      showValue ? formatValue(value, series, index) : undefined,
      showCategory ? category : undefined,
      showSeries ? series.name : undefined,
      showPercent && total ? formatChartNumberWithColor(value / total, series.valueNumberFormatOverrides[index] ?? series.numberFormat ?? chart.numberFormat ?? '0.0%', chart.date1904, chart.indexedColors) : undefined,
      showBubbleSize ? formatValue(bubbleSize, series, index) : undefined,
    ].filter((part): part is string | ReturnType<typeof formatChartNumberWithColor> => part !== undefined)
    if (!parts.length) return showLegendKey ? {text: ''} : undefined
    const defaultSeparator = isPieLike && !showValue && showCategory && !showSeries && showPercent ? '\n' : ','
    const color = parts.find((part): part is ReturnType<typeof formatChartNumberWithColor> => typeof part !== 'string' && Boolean(part.color))
    return {
      text: parts.map(part => typeof part === 'string' ? part : part.text).join(series.labelSeparatorOverrides[index] ?? series.labelSeparator ?? chart.labelSeparator ?? defaultSeparator),
      ...(color?.color ? {color: color.color} : {}),
    }
  }
  if (stock) {
    const highSeries = chart.series[stock.highIndex]
    const lowSeries = chart.series[stock.lowIndex]
    const closeSeries = chart.series[stock.closeIndex]
    const openSeries = stock.openIndex === undefined ? undefined : chart.series[stock.openIndex]
    const highLowSegments: string[] = []
    const dropSegments: string[] = []
    const priceIndices = [stock.openIndex, stock.highIndex, stock.lowIndex, stock.closeIndex].filter((index): index is number => index !== undefined)
    const bodyWidthFactor = 100 / (100 + stock.gapWidth)
    const volumeSeries = stock.volumeGapWidth === undefined || stock.volumeIndex === undefined ? undefined : chart.series[stock.volumeIndex]
    if (volumeSeries) {
      const {minimum: volumeMinimum, upper: volumeUpper} = axisBoundsForSeries(volumeSeries)
      const zeroY = seriesValuePosition(volumeSeries, clamp(0, volumeMinimum, volumeUpper))
      const volumeWidthFactor = 100 / (100 + (stock.volumeGapWidth ?? 150))
      for (let index = 0; index < chart.categories.length; index++) {
        const value = volumeSeries.values[index]
        if (value == null) continue
        const x = categoryX(index)
        const volumeWidth = categoryBandSize(index) * volumeWidthFactor
        const valueY = seriesValuePosition(volumeSeries, value)
        const top = Math.min(zeroY, valueY)
        const bottom = Math.max(zeroY, valueY)
        bars.push({x: x - volumeWidth / 2, y: top, width: volumeWidth, height: Math.max(1, bottom - top), color: volumeSeries.pointColors[index] || volumeSeries.color, stroke: volumeSeries.pointStrokes?.[index] ?? volumeSeries.stroke, strokeWidth: chartStrokeWidth(volumeSeries.pointStrokeWidthsEmu?.[index] ?? volumeSeries.strokeWidthEmu), ...chartStrokeStyle(volumeSeries, index, chartStrokeWidth(volumeSeries.pointStrokeWidthsEmu?.[index] ?? volumeSeries.strokeWidthEmu) ?? 1)})
        const label = formatDataLabel(volumeSeries, index, value, chart.categories[index] || String(index + 1))
        if (label !== undefined) {
          const position = labelPosition(volumeSeries, index, 'outEnd')
          let labelX = x
          let labelY = value >= 0 ? top - 8 : bottom + 20
          let anchor = 'middle'
          if (position === 'ctr') labelY = (zeroY + valueY) / 2 + 6
          else if (position === 'inBase') labelY = value >= 0 ? zeroY - 8 : zeroY + 18
          else if (position === 'inEnd') labelY = value >= 0 ? valueY + 18 : valueY - 4
          else if (position === 'b') labelY = bottom + 20
          else if (position === 'l') { labelX -= volumeWidth / 2 + 8; labelY = (zeroY + valueY) / 2 + 6; anchor = 'end' }
          else if (position === 'r') { labelX += volumeWidth / 2 + 8; labelY = (zeroY + valueY) / 2 + 6; anchor = 'start' }
          pushDataLabel(label, labelX, labelY, anchor, volumeSeries, index)
        }
      }
    }
    for (let index = 0; index < chart.categories.length; index++) {
      const highValue = highSeries?.values[index]
      const lowValue = lowSeries?.values[index]
      const closeValue = closeSeries?.values[index]
      const openValue = openSeries?.values[index]
      if (highValue == null || lowValue == null || closeValue == null || (openSeries && openValue == null)) continue
      const x = categoryX(index)
      const bodyWidth = categoryBandSize(index) * bodyWidthFactor
      const highY = valuePosition(Math.max(highValue, lowValue))
      const lowY = valuePosition(Math.min(highValue, lowValue))
      const closeY = valuePosition(closeValue)
      if (stock.showHighLowLines) highLowSegments.push(`M ${x} ${highY} L ${x} ${lowY}`)
      if (openSeries && openValue != null) {
        const openY = valuePosition(openValue)
        if (stock.showUpDownBars) {
          const color = closeValue >= openValue ? stock.upColor : stock.downColor
          bars.push({x: x - bodyWidth / 2, y: Math.min(openY, closeY), width: bodyWidth, height: Math.max(1, Math.abs(closeY - openY)), color})
        }
        highLowSegments.push(`M ${x - 7} ${openY} L ${x} ${openY}`, `M ${x} ${closeY} L ${x + 7} ${closeY}`)
      } else {
        highLowSegments.push(`M ${x - 7} ${closeY} L ${x + 7} ${closeY}`)
      }
      if (stock.showDropLines) dropSegments.push(`M ${x} ${closeY} L ${x} ${categoryAxisCrossingCoordinate}`)
      const label = formatDataLabel(closeSeries!, index, closeValue, chart.categories[index] || String(index + 1))
      if (label !== undefined) {
        const position = labelPosition(closeSeries!, index, 'r')
        const labelX = position === 'l' ? x - 10 : position === 'r' ? x + 10 : x
        const labelY = position === 't' ? closeY - 12 : position === 'b' ? closeY + 20 : closeY + 5
        pushDataLabel(label, labelX, labelY, position === 'l' ? 'end' : position === 'r' ? 'start' : 'middle', closeSeries!, index)
      }
    }
    if (highLowSegments.length) lines.push({d: highLowSegments.join(' '), color: stock.highLowColor, width: 2})
    if (dropSegments.length) lines.push({d: dropSegments.join(' '), color: '#94a3b8', width: 1})
    if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== 'none') {
      const categoryLabelY = categoryTickLabelPosition === 'high'
        ? chart.valueAxis?.reverse ? plot.top + plot.height + 25 : plot.top - 10
        : categoryTickLabelPosition === 'low'
          ? chart.valueAxis?.reverse ? plot.top - 10 : plot.top + plot.height + 25
          : categoryAxisSide === 't' ? categoryAxisCrossingCoordinate - 10 : categoryAxisCrossingCoordinate + 25
      chart.categories.forEach((text, index) => {
        if (showCategoryAxisLabel(index)) categories.push({...formatCategoryAxisLabel(text), x: categoryX(index), y: categoryLabelY, anchor: 'middle'})
      })
    }
    for (const index of [...(volumeSeries && stock.volumeIndex !== undefined ? [stock.volumeIndex] : []), ...priceIndices]) {
      const series = chart.series[index]
      if (series && !chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({text: series.name, x: 88 + legend.length * 220, y: 530, ...chartLegendKey(series)})
    }
  } else if (chart.type === 'area' && !combo) {
    chart.series.forEach((series, seriesIndex) => {
      let upperPoints: Array<{x: number; y: number}> = []
      let lowerPoints: Array<{x: number; y: number}> = []
      const flushArea = () => {
        if (upperPoints.length > 1) {
          const points = [...upperPoints, ...lowerPoints.reverse()]
          areas.push({d: `${points.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')} Z`, color: series.color, stroke: series.stroke ?? series.color, strokeWidth: chartStrokeWidth(series.strokeWidthEmu), ...chartStrokeStyle(series, undefined, chartStrokeWidth(series.strokeWidthEmu) ?? 1)})
        }
        upperPoints = []
        lowerPoints = []
      }
      for (let index = 0; index < categoryCount; index++) {
        const value = series.values[index]
        if (value === null || value === undefined) {
          flushArea()
          continue
        }
        const x = categoryX(index)
        const segment = stacking?.segments[seriesIndex]?.[index]
        const upperValue = stacking ? segment?.end ?? value : value
        const lowerValue = stacking ? segment?.start ?? 0 : categoryAxisCrossingValue
        const upperY = valuePosition(upperValue)
        const lowerY = valuePosition(lowerValue)
        upperPoints.push({x, y: upperY})
        lowerPoints.push({x, y: lowerY})
        const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
        if (label !== undefined) {
          const position = labelPosition(series, index, stacking ? 'ctr' : 'outEnd')
          const labelY = position === 'ctr' ? (upperY + lowerY) / 2 + 5
            : position === 'inBase' ? lowerY - 8
              : position === 'inEnd' ? upperY + 18
                : position === 'b' ? Math.max(upperY, lowerY) + 20
                  : Math.min(upperY, lowerY) - 8
          pushDataLabel(label, x, labelY, 'middle', series, index)
        }
      }
      flushArea()
      if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({text: series.name, x: 88 + (legend.length % 4) * 220, y: 530 + Math.floor(legend.length / 4) * 22, ...chartLegendKey(series)})
    })
  }
  else if (bubble) {
    const sizes = chart.series.flatMap(series => (series.bubbleSizes || []).filter((size): size is number => size !== null && size !== undefined && size !== 0 && (size > 0 || chart.showNegativeBubbles === true)))
    const maximumSize = sizes.reduce((maximum, size) => Math.max(maximum, Math.abs(size)), 0)
    // ponytail: normalized bubbles top out at 24% of the plot's short side; native-size calibration needs reference screenshots.
    const maximumRadius = Math.min(plot.width, plot.height) * 0.08 * (chart.bubbleScale ?? 100) / 100
    chart.series.forEach((series, seriesIndex) => {
      const pointCount = Math.max(series.xValues?.length || 0, series.values.length, series.bubbleSizes?.length || 0)
      for (let index = 0; index < pointCount; index++) {
        const xValue = series.xValues?.[index]
        const yValue = series.values[index]
        const bubbleSize = series.bubbleSizes?.[index]
        if (xValue == null || yValue == null || bubbleSize == null || bubbleSize === 0 || (bubbleSize < 0 && chart.showNegativeBubbles !== true) || maximumSize <= 0) continue
        const fraction = Math.abs(bubbleSize) / maximumSize
        const radius = maximumRadius * (chart.bubbleSizeRepresents === 'w' ? fraction : Math.sqrt(fraction))
        if (radius <= 0) continue
        const x = xPosition(xValue)
        const y = valuePosition(yValue)
        const color = series.pointColors[index] || (chart.bubbleVaryColors ? chart.palette[index % chart.palette.length] || series.color : series.color)
        markers.push({ x, y, color, stroke: series.pointStrokes?.[index] ?? series.stroke, strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu), ...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1), radius })
        const label = formatDataLabel(series, index, yValue, String(xValue), undefined, bubbleSize)
        if (label !== undefined) {
          const labelType = labelPosition(series, index, 'r')
          const labelX = labelType === 'l' ? x - radius - 4 : labelType === 'r' ? x + radius + 4 : x
          const labelY = labelType === 't' ? y - radius - 4 : labelType === 'b' ? y + radius + 18 : y + 5
          pushDataLabel(label, labelX, labelY, labelType === 'l' ? 'end' : labelType === 'r' ? 'start' : 'middle', series, index)
        }
      }
      if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({ text: series.name, x: 88 + seriesIndex * 220, y: 530, ...chartLegendKey(series) })
    })
  } else if (scatter) {
    chart.series.forEach((series, seriesIndex) => {
      let path = ''
      let segment: Array<{x: number; y: number}> = []
      const flushSegment = () => {
        path += smoothChartLinePath(segment, series.scatterSmooth === true)
        segment = []
      }
      const pointCount = Math.max(series.xValues?.length || 0, series.values.length)
      for (let index = 0; index < pointCount; index++) {
        const xValue = series.xValues?.[index]
        const yValue = series.values[index]
        if (xValue == null || yValue == null) {
          flushSegment()
          continue
        }
        const x = xPosition(xValue)
        const y = valuePosition(yValue)
        if (series.scatterShowLine) segment.push({x, y})
        const pointMarkerSymbol = series.markerSymbolOverrides?.[index] ?? series.markerSymbol
        const showMarker = pointMarkerSymbol === undefined ? series.scatterShowMarker : pointMarkerSymbol !== 'none'
        if (showMarker) markers.push({
          x,
          y,
          color: series.pointColors[index] || series.color,
          stroke: series.pointStrokes?.[index] ?? series.stroke,
          strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
          ...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
          symbol: pointMarkerSymbol ?? 'circle',
          radius: series.markerSizeOverrides?.[index] ?? series.markerSize,
        })
        const label = formatDataLabel(series, index, yValue, String(xValue))
        if (label !== undefined) {
          const labelType = labelPosition(series, index, 'r')
          const labelX = labelType === 'l' ? x - 10 : labelType === 'r' ? x + 10 : x
          const labelY = labelType === 't' ? y - 12 : labelType === 'b' ? y + 24 : y + 5
          pushDataLabel(label, labelX, labelY, labelType === 'l' ? 'end' : labelType === 'r' ? 'start' : 'middle', series, index)
        }
      }
      flushSegment()
      if (path) lines.push({ d: path, color: series.stroke ?? series.color, width: chartStrokeWidth(series.strokeWidthEmu), ...chartStrokeStyle(series, undefined, chartStrokeWidth(series.strokeWidthEmu) ?? 4) })
      if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({ text: series.name, x: 88 + seriesIndex * 220, y: 530, ...chartLegendKey(series) })
    })
  } else if (chart.type === 'pie' || chart.type === 'doughnut') {
    const centerX = plot.left + plot.width / 2
    const centerY = plot.top + plot.height / 2
    const radius = Math.min(plot.width, plot.height) * 0.47
    const doughnut = chart.type === 'doughnut'
    const baseInnerRadius = doughnut ? radius * (chart.holeSize ?? 50) / 100 : 0
    const ringSeries = doughnut ? chart.series : chart.series.slice(0, 1)
    const ringWidth = doughnut ? (radius - baseInnerRadius) / Math.max(ringSeries.length, 1) : radius
    ringSeries.forEach((series, seriesIndex) => {
      const innerRadius = doughnut ? baseInnerRadius + ringWidth * seriesIndex : 0
      const outerRadius = doughnut ? innerRadius + ringWidth : radius
      const points = chart.categories.flatMap((text, index) => {
        const value = series.values[index]
        return value != null && value > 0 ? [{ text, index, value, color: series.pointColors[index] || chart.palette[index % chart.palette.length] || series.color }] : []
      })
      const total = points.reduce((sum, point) => sum + point.value, 0)
      let angle = -Math.PI / 2 + (chart.firstSliceAngle ?? 0) * Math.PI / 180
      for (const point of points) {
        const sweep = point.value / total * Math.PI * 2
        const startX = centerX + Math.cos(angle) * outerRadius
        const startY = centerY + Math.sin(angle) * outerRadius
        const endAngle = angle + sweep
        const endX = centerX + Math.cos(endAngle) * outerRadius
        const endY = centerY + Math.sin(endAngle) * outerRadius
        const endInnerX = centerX + Math.cos(endAngle) * innerRadius
        const endInnerY = centerY + Math.sin(endAngle) * innerRadius
        const startInnerX = centerX + Math.cos(angle) * innerRadius
        const startInnerY = centerY + Math.sin(angle) * innerRadius
        const fullCircle = sweep >= Math.PI * 2 - 0.0001
        const d = innerRadius > 0
          ? fullCircle
            ? `M ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 1 1 ${centerX - Math.cos(angle) * outerRadius} ${centerY - Math.sin(angle) * outerRadius} A ${outerRadius} ${outerRadius} 0 1 1 ${startX} ${startY} L ${startInnerX} ${startInnerY} A ${innerRadius} ${innerRadius} 0 1 0 ${centerX - Math.cos(angle) * innerRadius} ${centerY - Math.sin(angle) * innerRadius} A ${innerRadius} ${innerRadius} 0 1 0 ${startInnerX} ${startInnerY} Z`
            : `M ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${endX} ${endY} L ${endInnerX} ${endInnerY} A ${innerRadius} ${innerRadius} 0 ${sweep > Math.PI ? 1 : 0} 0 ${startInnerX} ${startInnerY} Z`
          : fullCircle
            ? `M ${centerX} ${centerY} L ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 1 1 ${centerX - Math.cos(angle) * outerRadius} ${centerY - Math.sin(angle) * outerRadius} A ${outerRadius} ${outerRadius} 0 1 1 ${startX} ${startY} Z`
            : `M ${centerX} ${centerY} L ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${endX} ${endY} Z`
        const strokeWidth = chartStrokeWidth(series.pointStrokeWidthsEmu?.[point.index] ?? series.strokeWidthEmu)
        slices.push({
          d,
          color: point.color,
          stroke: series.pointStrokes?.[point.index] ?? series.stroke ?? 'white',
          strokeWidth: strokeWidth ?? 2,
          ...chartStrokeStyle(series, point.index, strokeWidth ?? 2),
        })
        const label = formatDataLabel(series, point.index, point.value, point.text, total)
        if (label !== undefined) {
          const labelAngle = angle + sweep / 2
          const position = labelPosition(series, point.index, 'ctr')
          const showLeaderLine = series.leaderLineOverrides?.[point.index] ?? series.showLeaderLines ?? false
          const hasLeaderLine = showLeaderLine && position === 'outEnd'
          const labelRadius = doughnut
            ? position === 'outEnd' ? outerRadius * 1.12
              : position === 'inEnd' ? innerRadius + ringWidth * 0.82
                : position === 'inBase' ? innerRadius + ringWidth * 0.18
                  : (outerRadius + innerRadius) / 2
            : radius * (position === 'outEnd' ? 1.12 : position === 'inEnd' ? 0.82 : position === 'inBase' ? 0.24 : position === 'bestFit' ? 0.62 : 0.68)
          const side = Math.cos(labelAngle) >= 0 ? 1 : -1
          const leaderY = centerY + Math.sin(labelAngle) * outerRadius * 1.04
          const x = hasLeaderLine ? centerX + side * outerRadius * 1.12 : position === 'l' ? centerX - radius * 0.75 : position === 'r' ? centerX + radius * 0.75 : centerX + Math.cos(labelAngle) * labelRadius
          const y = hasLeaderLine ? leaderY + 6 : position === 't' ? centerY - radius * 0.78 : position === 'b' ? centerY + radius * 0.78 : centerY + Math.sin(labelAngle) * labelRadius + 6
          const anchor = hasLeaderLine ? side > 0 ? 'start' : 'end' : position === 'l' ? 'end' : position === 'r' ? 'start' : 'middle'
          const leaderLine = hasLeaderLine
            ? `M ${centerX + Math.cos(labelAngle) * outerRadius} ${centerY + Math.sin(labelAngle) * outerRadius} L ${centerX + Math.cos(labelAngle) * outerRadius * 1.04} ${leaderY} L ${x - side * 6} ${leaderY}`
            : undefined
          pushDataLabel(label, x, y, anchor, series, point.index, leaderLine)
        }
        angle = endAngle
      }
      if (seriesIndex === 0) points.slice(0, 12).forEach(point => {
        if (chart.legend?.hiddenEntries?.includes(point.index)) return
        legend.push({ text: point.text, x: 590, y: 105 + legend.length * 28, color: point.color })
      })
    })
  } else if (radar) {
    const radialValues = ticks.map(tick => axisFraction(tick.value)).filter(fraction => fraction > 0.0001 && fraction <= 1.0001)
    if (!chart.valueAxis?.deleted && chart.valueAxis?.majorGridlines) {
      for (const fraction of [...new Set([...radialValues, 1])].sort((a, b) => a - b)) {
        const points = chart.categories.map((_category, index) => radarPoint(index, fraction))
        radarGrid.push({ d: `${points.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')} Z` })
      }
    }
    if (!chart.categoryAxisDeleted && chart.categoryAxisMajorGridlines) {
      chart.categories.forEach((_category, index) => {
        if (!showCategoryAxisTick(index)) return
        const point = radarPoint(index)
        radarSpokes.push({ x1: radarCenterX, y1: radarCenterY, x2: point.x, y2: point.y })
      })
    }
    if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== 'none') {
      chart.categories.forEach((text, index) => {
        if (!showCategoryAxisLabel(index)) return
        const point = radarPoint(index, 1, 22)
        const horizontal = Math.cos(point.angle)
        categories.push({
          ...formatCategoryAxisLabel(text),
          x: point.x,
          y: point.y + (Math.sin(point.angle) > 0.55 ? 16 : Math.sin(point.angle) < -0.55 ? -8 : 5),
          anchor: horizontal > 0.25 ? 'start' : horizontal < -0.25 ? 'end' : 'middle',
        })
      })
    }
    chart.series.forEach(series => {
      const points = chart.categories.map((_category, index) => {
        const value = series.values[index]
        return value == null ? null : radarPoint(index, axisFraction(value))
      })
      const complete = points.length >= 3 && points.every((point): point is NonNullable<typeof point> => point !== null)
      if (chart.radarStyle === 'filled' && complete) {
        areas.push({
          d: `${points.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')} Z`,
          color: series.color,
          stroke: series.stroke ?? series.color,
          strokeWidth: chartStrokeWidth(series.strokeWidthEmu),
          ...chartStrokeStyle(series, undefined, chartStrokeWidth(series.strokeWidthEmu) ?? 1),
          opacity: 0.22,
        })
      }
      if (series.radarShowLine !== false) {
        const path = complete
          ? `${points.map((point, index) => `${index ? 'L' : 'M'} ${point!.x} ${point!.y}`).join(' ')} Z`
          : points.flatMap((point, index) => {
              const next = points[(index + 1) % points.length]
              return point && next ? [`M ${point.x} ${point.y} L ${next.x} ${next.y}`] : []
            }).join(' ')
        if (path) lines.push({ d: path, color: series.stroke ?? series.color, width: chartStrokeWidth(series.strokeWidthEmu), ...chartStrokeStyle(series, undefined, chartStrokeWidth(series.strokeWidthEmu) ?? 4) })
      }
      points.forEach((point, index) => {
        if (!point) return
        const pointMarkerSymbol = series.markerSymbolOverrides?.[index] ?? series.markerSymbol
        const showMarker = pointMarkerSymbol === undefined ? series.radarShowMarker : pointMarkerSymbol !== 'none'
        if (showMarker) markers.push({
          x: point.x,
          y: point.y,
          color: series.pointColors[index] || series.color,
          stroke: series.pointStrokes?.[index] ?? series.stroke,
          strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
          ...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
          symbol: pointMarkerSymbol ?? 'circle',
          radius: series.markerSizeOverrides?.[index] ?? series.markerSize ?? series.radarMarkerSize ?? 5,
        })
        const value = series.values[index]
        if (value == null) return
        const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
        if (label === undefined) return
        const position = labelPosition(series, index, 'outEnd')
        const offset = position === 'inBase' || position === 'inEnd' || position === 'ctr' ? -10 : 14
        const x = position === 'l' ? point.x - 12 : position === 'r' ? point.x + 12 : point.x + Math.cos(point.angle) * offset
        const y = position === 't' ? point.y - 14 : position === 'b' ? point.y + 22 : point.y + Math.sin(point.angle) * offset + 6
        const anchor = position === 'l' ? 'end' : position === 'r' ? 'start' : Math.cos(point.angle) > 0.25 ? 'start' : Math.cos(point.angle) < -0.25 ? 'end' : 'middle'
        pushDataLabel(label, x, y, anchor, series, index)
      })
      if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({ text: series.name, x: 88 + legend.length * 220, y: 530, ...chartLegendKey(series) })
    })
  } else {
    const barSeriesByGroup = new Map<number, number[]>()
    chart.series.forEach((series, seriesIndex) => {
      if (seriesType(series) !== 'bar') return
      const groupIndex = series.chartGroupIndex ?? 0
      barSeriesByGroup.set(groupIndex, [...(barSeriesByGroup.get(groupIndex) || []), seriesIndex])
    })
    const linePointsBySeries = new Map<number, Array<{x: number; y: number} | null>>()
    chart.series.forEach((series, seriesIndex) => {
      if (seriesType(series) === 'line' && series.lineShowLine !== false) linePointsBySeries.set(seriesIndex, [])
    })
    for (let index = 0; index < categoryCount; index++) {
      const centerX = categoryX(index)
      const centerY = categoryY(index)
      if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== 'none') {
        const categoryLabel = categoryTickLabelPosition === 'high'
          ? horizontalValueAxis
            ? { x: chart.valueAxis?.reverse ? plot.left - 10 : plot.left + plot.width + 10, y: centerY + 5, anchor: chart.valueAxis?.reverse ? 'end' : 'start' }
            : { x: centerX, y: chart.valueAxis?.reverse ? plot.top + plot.height + 25 : plot.top - 10, anchor: 'middle' }
          : categoryTickLabelPosition === 'low'
            ? horizontalValueAxis
              ? { x: chart.valueAxis?.reverse ? plot.left + plot.width + 10 : plot.left - 10, y: centerY + 5, anchor: chart.valueAxis?.reverse ? 'start' : 'end' }
              : { x: centerX, y: chart.valueAxis?.reverse ? plot.top - 10 : plot.top + plot.height + 25, anchor: 'middle' }
            : categoryAxisSide === 'l'
              ? { x: categoryAxisCrossingCoordinate - 10, y: centerY + 5, anchor: 'end' }
              : categoryAxisSide === 'r'
                ? { x: categoryAxisCrossingCoordinate + 10, y: centerY + 5, anchor: 'start' }
                : categoryAxisSide === 't'
                  ? { x: centerX, y: categoryAxisCrossingCoordinate - 10, anchor: 'middle' }
                  : { x: centerX, y: categoryAxisCrossingCoordinate + 25, anchor: 'middle' }
        if (showCategoryAxisLabel(index)) categories.push({ ...formatCategoryAxisLabel(chart.categories[index] || ''), ...categoryLabel })
      }
      chart.series.forEach((series, seriesIndex) => {
        const value = series.values[index]
        if (value === null || value === undefined) {
          linePointsBySeries.get(seriesIndex)?.push(null)
          return
        }
        const color = series.color
        const kind = seriesType(series)
        const stackedSegment = stackedSegmentFor(seriesIndex)?.[index]
        const grouping = series.grouping || chart.grouping || 'clustered'
        const isStacked = grouping !== 'clustered'
        if (kind === 'line') {
          const x = horizontalValueAxis ? seriesValuePosition(series, stackedSegment?.end ?? value) : centerX
          const y = horizontalValueAxis ? centerY : seriesValuePosition(series, stackedSegment?.end ?? value)
          linePointsBySeries.get(seriesIndex)?.push({x, y})
          const pointMarkerSymbol = series.markerSymbolOverrides?.[index] ?? series.markerSymbol
          const showMarker = pointMarkerSymbol === undefined ? series.lineShowMarker !== false : pointMarkerSymbol !== 'none'
          if (showMarker) markers.push({
            x,
            y,
            color: series.pointColors[index] || color,
            stroke: series.pointStrokes?.[index] ?? series.stroke,
            strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
            ...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
            symbol: pointMarkerSymbol ?? 'circle',
            radius: series.markerSizeOverrides?.[index] ?? series.markerSize,
          })
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, 't')
            const labelX = position === 'l' ? x - 10 : position === 'r' ? x + 10 : x
            const labelY = position === 'b' ? y + 24 : position === 'ctr' || position === 'inBase' || position === 'inEnd' ? y + 7 : y - 12
            const anchor = position === 'l' ? 'end' : position === 'r' ? 'start' : 'middle'
            pushDataLabel(label, labelX, labelY, anchor, series, index)
          }
        } else if (kind === 'area') {
          const x = horizontalValueAxis ? seriesValuePosition(series, stackedSegment?.end ?? value) : centerX
          const y = horizontalValueAxis ? centerY : seriesValuePosition(series, stackedSegment?.end ?? value)
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, 't')
            const labelX = position === 'l' ? x - 10 : position === 'r' ? x + 10 : x
            const labelY = position === 'b' ? y + 24 : position === 'ctr' || position === 'inBase' || position === 'inEnd' ? y + 7 : y - 12
            const anchor = position === 'l' ? 'end' : position === 'r' ? 'start' : 'middle'
            pushDataLabel(label, labelX, labelY, anchor, series, index)
          }
        } else if (chart.direction === 'horizontal') {
          const groupIndexes = barSeriesByGroup.get(series.chartGroupIndex ?? 0) || [seriesIndex]
          const barOrdinal = groupIndexes.indexOf(seriesIndex)
          const seriesCount = Math.max(groupIndexes.length, 1)
          const {minimum: axisMinimum, upper: axisUpper} = axisBoundsForSeries(series)
          const groupHeight = categoryBandSize(index) * 0.72
          const barHeight = groupHeight / (isStacked ? 1 : seriesCount)
          const barY = centerY - groupHeight / 2 + (isStacked ? 0 : barOrdinal * barHeight)
          const zeroX = seriesValuePosition(series, Math.max(axisMinimum, Math.min(axisUpper, 0)))
          const baseX = isStacked ? seriesValuePosition(series, stackedSegment?.start ?? 0) : zeroX
          const valueX = seriesValuePosition(series, isStacked ? stackedSegment?.end ?? value : value)
          bars.push({ x: Math.min(baseX, valueX), y: barY, width: Math.max(1, Math.abs(valueX - baseX)), height: barHeight * 0.86, color: series.pointColors[index] || color, stroke: series.pointStrokes?.[index] ?? series.stroke, strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu), ...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1) })
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, isStacked ? 'ctr' : 'outEnd')
            const direction = value >= 0 ? 1 : -1
            const left = Math.min(baseX, valueX)
            const right = Math.max(baseX, valueX)
            let labelX = valueX + direction * 8
            let labelY = barY + barHeight * 0.68
            let anchor = direction > 0 ? 'start' : 'end'
            if (position === 'ctr') { labelX = (baseX + valueX) / 2; anchor = 'middle' }
            else if (position === 'inBase') { labelX = baseX + direction * 20; anchor = direction > 0 ? 'start' : 'end' }
            else if (position === 'inEnd') { labelX = valueX - direction * 20; anchor = direction > 0 ? 'end' : 'start' }
            else if (position === 'l') { labelX = left - 8; anchor = 'end' }
            else if (position === 'r') { labelX = right + 8; anchor = 'start' }
            else if (position === 't') { labelX = (zeroX + valueX) / 2; labelY = barY - 8; anchor = 'middle' }
            else if (position === 'b') { labelX = (zeroX + valueX) / 2; labelY = barY + barHeight + 15; anchor = 'middle' }
            pushDataLabel(label, labelX, labelY, anchor, series, index)
          }
        } else {
          const groupIndexes = barSeriesByGroup.get(series.chartGroupIndex ?? 0) || [seriesIndex]
          const barOrdinal = groupIndexes.indexOf(seriesIndex)
          const seriesCount = Math.max(groupIndexes.length, 1)
          const {minimum: axisMinimum, upper: axisUpper} = axisBoundsForSeries(series)
          const groupWidth = categoryBandSize(index) * 0.72
          const barWidth = groupWidth / (isStacked ? 1 : seriesCount)
          const zeroY = seriesValuePosition(series, Math.max(axisMinimum, Math.min(axisUpper, 0)))
          const baseY = isStacked ? seriesValuePosition(series, stackedSegment?.start ?? 0) : zeroY
          const valueY = seriesValuePosition(series, isStacked ? stackedSegment?.end ?? value : value)
          const barX = centerX - groupWidth / 2 + (isStacked ? 0 : barOrdinal * barWidth)
          bars.push({ x: barX, y: Math.min(baseY, valueY), width: barWidth * 0.86, height: Math.max(1, Math.abs(valueY - baseY)), color: series.pointColors[index] || color, stroke: series.pointStrokes?.[index] ?? series.stroke, strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu), ...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1) })
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, isStacked ? 'ctr' : 'outEnd')
            const top = Math.min(baseY, valueY)
            const bottom = Math.max(baseY, valueY)
            let labelX = barX + barWidth * 0.43
            let labelY = value >= 0 ? top - 8 : bottom + 20
            let anchor = 'middle'
            if (position === 'ctr') labelY = (baseY + valueY) / 2 + 6
            else if (position === 'inBase') labelY = value >= 0 ? baseY - 8 : baseY + 18
            else if (position === 'inEnd') labelY = value >= 0 ? valueY + 18 : valueY - 4
            else if (position === 'b') labelY = bottom + 20
            else if (position === 'l') { labelX = barX - 8; labelY = (baseY + valueY) / 2 + 6; anchor = 'end' }
            else if (position === 'r') { labelX = barX + barWidth * 0.86 + 8; labelY = (baseY + valueY) / 2 + 6; anchor = 'start' }
            pushDataLabel(label, labelX, labelY, anchor, series, index)
          }
        }
      })
    }
    linePointsBySeries.forEach((points, seriesIndex) => {
      const series = chart.series[seriesIndex]
      if (!series) return
      let path = ''
      let segment: Array<{x: number; y: number}> = []
      const flushSegment = () => {
        path += smoothChartLinePath(segment, series.lineSmooth === true)
        segment = []
      }
      points.forEach(point => {
        if (point) segment.push(point)
        else flushSegment()
      })
      flushSegment()
      if (path) lines.push({ d: path, color: series.stroke ?? series.color, width: chartStrokeWidth(series.strokeWidthEmu), ...chartStrokeStyle(series, undefined, chartStrokeWidth(series.strokeWidthEmu) ?? 4) })
    })
    chart.series.slice(0, 12).forEach(series => {
      if (chart.legend?.hiddenEntries?.includes(series.legendIndex)) return
      legend.push({ text: series.name, x: 88 + (legend.length % 4) * 220, y: 530 + Math.floor(legend.length / 4) * 22, ...chartLegendKey(series) })
    })
  }

  if (!chart.legend) legend.length = 0
  else {
    const pos = chart.legend!.position
    const defaultX = defaultLegendOrigin(pos, 'x')
    const defaultY = defaultLegendOrigin(pos, 'y')
    const manual = chart.legend!.manualLayout
    const originX = manual?.x ? (manual.x.mode === 'edge' ? manual.x.value * 1000 : defaultX + manual.x.value * 1000) : defaultX
    const originY = manual?.y ? (manual.y.mode === 'edge' ? manual.y.value * 600 : defaultY + manual.y.value * 600) : defaultY
    const offsetX = originX - defaultX
    const offsetY = originY - defaultY
    legend.forEach((entry, index) => {
      const x = pos === 'left' ? 88 : pos === 'right' ? 780 : pos === 'topRight' ? 760 : 88 + index % 4 * 220
      const y = pos === 'left' || pos === 'right' ? 105 + index * 28
      : pos === 'top' ? 77 + Math.floor(index / 4) * 18
        : pos === 'topRight' ? 92 + index * 22
          : 530 + Math.floor(index / 4) * 22
      legend[index] = { ...entry, x: x + offsetX, y: y + offsetY }
    })
  }
  const manualLayout = chart.legend?.manualLayout
  let legendClip: ChartVisual['legendClip']
  if (manualLayout?.w || manualLayout?.h) {
    const x = manualLayout.x?.mode === 'edge' ? manualLayout.x.value * 1000 : defaultLegendOrigin(chart.legend?.position, 'x') + (manualLayout.x?.value || 0) * 1000
    const y = manualLayout.y?.mode === 'edge' ? manualLayout.y.value * 600 : defaultLegendOrigin(chart.legend?.position, 'y') + (manualLayout.y?.value || 0) * 600
    const width = manualLayout.w?.mode === 'edge' ? manualLayout.w.value * 1000 - x : (manualLayout.w?.value ?? 1) * 1000
    const height = manualLayout.h?.mode === 'edge' ? manualLayout.h.value * 600 - y : (manualLayout.h?.value ?? 1) * 600
    legendClip = {
      id: `${customGradientScope}-legend-${elementId}`.replace(/[^a-zA-Z0-9_-]/g, '_'),
      x,
      y,
      width: Math.max(0, width),
      height: Math.max(0, height),
    }
  }
  const baseline = xyChart
    ? { x1: plot.left, y1: xyXAxisCrossingY, x2: plot.left + plot.width, y2: xyXAxisCrossingY }
    : horizontalValueAxis
      ? { x1: categoryAxisCrossingCoordinate, y1: plot.top, x2: categoryAxisCrossingCoordinate, y2: plot.top + plot.height }
      : { x1: plot.left, y1: categoryAxisCrossingCoordinate, x2: plot.left + plot.width, y2: categoryAxisCrossingCoordinate }
  if (dateAxisMajorTickValues?.length && categories.length && hasDateAxisScale) {
    const firstLabel = categories[0]!
    categories.splice(0, categories.length, ...dateAxisMajorTickValues.map(value => {
      const fraction = dateAxisFraction(value)
      const formattedLabel = formatCategoryAxisLabel(String(value))
      return {
        ...firstLabel,
        ...formattedLabel,
        color: formattedLabel.color,
        ...(horizontalValueAxis ? { y: plot.top + plot.height * fraction } : { x: plot.left + plot.width * fraction }),
      }
    }))
  }
  return { plot, ticks, xTicks, axisTicks, secondaryValueAxes, radarGrid, radarSpokes, categoryGridlines, categories, bars, areas, lines, markers, slices, legend, ...(legendClip ? { legendClip } : {}), dataLabels, baseline, valueAxisLine: radar ? undefined : valueAxisLine }
}

const props = withDefaults(defineProps<{
  slide: PptxSlide
  width: number
  height: number
  number: number
  animationStep?: number
  showAnimations?: boolean
  showNarration?: boolean
  thumbnail?: boolean
  playbackPreview?: boolean
  triggerPlayback?: PptxTriggerPlayback[]
}>(), { animationStep: 0, showAnimations: true, showNarration: true, thumbnail: false, playbackPreview: false })
const emit = defineEmits<{
  'trigger-playback': [playback: PptxTriggerPlayback[]]
  'activate-hyperlink': [hyperlink: PptxHyperlink]
  'stop-all-audio': []
}>()

const backgroundImageSize = ref<{ sourceUrl: string; url: string; width: number; height: number; repeatX: number; repeatY: number }>()
const backgroundTileWarning = ref<LocaleMessageKey | ''>('')
watch(() => [props.slide.backgroundImage?.url, Boolean(props.slide.backgroundImage?.tile), props.slide.backgroundImage?.tile?.flip || 'none', props.slide.backgroundImage?.crop.left, props.slide.backgroundImage?.crop.top, props.slide.backgroundImage?.crop.right, props.slide.backgroundImage?.crop.bottom] as const, ([url, tiled, flip], _oldValue, onCleanup) => {
  backgroundImageSize.value = undefined
  backgroundTileWarning.value = ''
  if (!import.meta.client || !url || !tiled) return
  const crop = props.slide.backgroundImage?.crop ?? {left: 0, top: 0, right: 0, bottom: 0}
  const probe = new Image()
  let current = true
  let generatedUrl: string | undefined
  onCleanup(() => {
    current = false
    probe.onload = null
    probe.onerror = null
    if (generatedUrl) URL.revokeObjectURL(generatedUrl)
  })
  probe.onerror = () => {
    if (current) backgroundTileWarning.value = 'tiledBackgroundDecodeFailed'
  }
  probe.onload = () => {
    if (!current || probe.naturalWidth < 1 || probe.naturalHeight < 1) return
    const width = probe.naturalWidth
    const height = probe.naturalHeight
    const flipX = flip === 'x' || flip === 'xy'
    const flipY = flip === 'y' || flip === 'xy'
    const hasCrop = Object.values(crop).some(value => value !== 0)
    const sourceLeft = width * crop.left / 100_000
    const sourceTop = height * crop.top / 100_000
    const sourceWidth = width - sourceLeft - width * crop.right / 100_000
    const sourceHeight = height - sourceTop - height * crop.bottom / 100_000
    const tileWidth = Math.max(1, Math.round(sourceWidth))
    const tileHeight = Math.max(1, Math.round(sourceHeight))
    const imageSize = { sourceUrl: url, url, width: tileWidth, height: tileHeight, repeatX: flipX ? 2 : 1, repeatY: flipY ? 2 : 1 }
    const originalImageSize = { sourceUrl: url, url, width, height, repeatX: 1, repeatY: 1 }
    if (!hasCrop && !flipX && !flipY) {
      backgroundImageSize.value = imageSize
      return
    }
    // ponytail: cap synthesized flip patterns at 16 MP; downsample if real decks hit this ceiling.
    const patternWidth = tileWidth * imageSize.repeatX
    const patternHeight = tileHeight * imageSize.repeatY
    if (patternWidth > 32_767 || patternHeight > 32_767 || patternWidth * patternHeight > 16_777_216) {
      backgroundTileWarning.value = hasCrop ? 'tiledBackgroundCropTooLarge' : 'tiledBackgroundFlipTooLarge'
      backgroundImageSize.value = originalImageSize
      return
    }
    const canvas = document.createElement('canvas')
    canvas.width = patternWidth
    canvas.height = patternHeight
    const context = canvas.getContext('2d')
    if (!context) {
      backgroundTileWarning.value = hasCrop ? 'tiledBackgroundCropUnsupported' : 'tiledBackgroundFlipUnsupported'
      backgroundImageSize.value = originalImageSize
      return
    }
    const drawTile = (x: number, y: number, mirrorX: boolean, mirrorY: boolean) => {
      context.save()
      context.translate(x + (mirrorX ? tileWidth : 0), y + (mirrorY ? tileHeight : 0))
      context.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1)
      context.drawImage(probe, sourceLeft, sourceTop, sourceWidth, sourceHeight, 0, 0, tileWidth, tileHeight)
      context.restore()
    }
    drawTile(0, 0, false, false)
    if (flipX) drawTile(tileWidth, 0, true, false)
    if (flipY) drawTile(0, tileHeight, false, true)
    if (flipX && flipY) drawTile(tileWidth, tileHeight, true, true)
    canvas.toBlob(blob => {
      if (!current) return
      if (!blob) {
        backgroundTileWarning.value = hasCrop ? 'tiledBackgroundCropEncodeFailed' : 'tiledBackgroundFlipEncodeFailed'
        backgroundImageSize.value = originalImageSize
        return
      }
      generatedUrl = URL.createObjectURL(blob)
      backgroundImageSize.value = { ...imageSize, url: generatedUrl }
    }, 'image/png')
  }
  probe.src = url
}, { immediate: true })

const frameStyle = computed<CSSProperties>(() => ({ aspectRatio: `${props.width} / ${props.height}` }))

const backgroundStyle = computed<CSSProperties>(() => {
  const image = props.slide.backgroundImage
  if (!image) return { background: props.slide.background || '#fff' }
  const { left, top, right, bottom } = image.crop
  const tile = image.tile
  const alignments = { tl: [0, 0], t: [50, 0], tr: [100, 0], l: [0, 50], ctr: [50, 50], r: [100, 50], bl: [0, 100], b: [50, 100], br: [100, 100] } as const
  const naturalSize = backgroundImageSize.value?.sourceUrl === image.url ? backgroundImageSize.value : undefined
  const tileWidth = naturalSize && tile ? naturalSize.width * tile.scaleX : 0
  const tileHeight = naturalSize && tile ? naturalSize.height * tile.scaleY : 0
  const tilePosition = (alignment: number, offset: number) => offset === 0
    ? `${alignment}%`
    : `calc(${alignment}% ${offset < 0 ? '-' : '+'} ${Math.abs(offset)}px)`
  const position = tile && tile.align === 'tl' && tile.offsetX === 0 && tile.offsetY === 0 && (naturalSize?.repeatX ?? 1) === 1 && (naturalSize?.repeatY ?? 1) === 1
    ? '0px 0px'
    : tile ? `${tilePosition(alignments[tile.align][0], tile.offsetX * scale.value + tileWidth * alignments[tile.align][0] / 100 * ((naturalSize?.repeatX ?? 1) - 1))} ${tilePosition(alignments[tile.align][1], tile.offsetY * scale.value + tileHeight * alignments[tile.align][1] / 100 * ((naturalSize?.repeatY ?? 1) - 1))}`
      : `${left + right ? left / (left + right) * 100 : 50}% ${top + bottom ? top / (top + bottom) * 100 : 50}%`
  return {
    background: props.slide.background || '#fff',
    backgroundImage: `url("${naturalSize?.url || image.url}")`,
    backgroundRepeat: tile ? 'repeat' : 'no-repeat',
    backgroundSize: tile
      ? naturalSize ? `${tileWidth * naturalSize.repeatX}px ${tileHeight * naturalSize.repeatY}px` : 'auto'
      : `${100 / (1 - (left + right) / 100000)}% ${100 / (1 - (top + bottom) / 100000)}%`,
    backgroundPosition: position,
  }
})

const stage = ref<HTMLElement | null>(null)
const stageWidth = ref(0)
const autoFitDimensions = reactive(new Map<string, { width: number; height: number }>())
const normalAutoFitScales = reactive(new Map<string, number>())
const textEllipsisLineCounts = reactive(new Map<string, number>())
const textAnchorCenterOffsets = new Map<string, number>()
const elementLayoutVersion = ref(0)
const textOverflowWarning = ref<LocaleMessageKey | ''>('')
const animationClock = ref(0)
const animationClockTargetIds = new Set<string>()
const animationTimelineVersion = ref(0)
const reduceMotion = ref(false)
const invalidMediaTrimIds = ref<string[]>([])
const invalidMediaFadeIds = ref<string[]>([])
const blockedMediaPlaybackIds = ref<string[]>([])
const mediaPlaybackFrames = new Map<HTMLMediaElement, number>()
const mediaFadeStates = new WeakMap<HTMLMediaElement, { baseVolume: number; appliedVolume: number; appliedFactor: number; baseMuted: boolean; appliedMuted: boolean }>()
const lastMediaBookmarkTimes = new WeakMap<HTMLMediaElement, number>()
const mediaBookmarkPlaybackStarted = new WeakSet<HTMLMediaElement>()
const mediaBookmarkTriggersByMedia = computed(() => {
  const triggersByMedia = new Map<string, Array<{ sequenceIndex: number; conditionIndex: number; timeMs: number }>>()
  props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
    sequence.conditions.forEach((condition, conditionIndex) => {
      if (condition.type !== 'mediaBookmark') return
      const triggers = triggersByMedia.get(condition.id) || []
      triggers.push({ sequenceIndex, conditionIndex, timeMs: condition.timeMs })
      triggersByMedia.set(condition.id, triggers)
    })
  })
  return triggersByMedia
})
const executedMediaActions = new Set<number>()
const scheduledMediaActions = new Map<string, number>()
const startedMediaActions = new Map<string, number>()
const completedMediaActions = reactive(new Map<string, number>())
const endedMediaTargets = reactive(new Map<string, number>())
const mediaActionEndTimers = new Map<string, ReturnType<typeof setTimeout>>()
const exclusivePausedMedia = new Map<HTMLMediaElement, {remainingMs?: number; actionKey?: string; pausedAt: number}>()
let automaticStartedAt: number | undefined
let stepStartedAt = new Map<number, number>()
let mainStepEndedAt = new Map<number, number>()
let triggerSteps = new Map<number, number>()
let triggerStartedAt = new Map<string, number>()
let triggerStepEndedAt = new Map<string, number>()
let triggerPlaybackByKey = new Map<string, PptxTriggerPlayback>()
const firedTimingNodeTriggers = new Set<string>()
const triggerConditionDueAt = new Map<string, number>()
const pendingTriggerConditions = new Set<string>()
const triggerConditionKey = (sequenceIndex: number, conditionIndex: number) => `${sequenceIndex}:${conditionIndex}`
let observer: ResizeObserver | undefined
let autoFitFrame = 0
let animationFrame = 0
let animationTimer: ReturnType<typeof setTimeout> | undefined
let motionPreference: MediaQueryList | undefined
let pixelRatioQuery: MediaQueryList | undefined
const pixelRatio = ref(1)
const scale = computed(() => stageWidth.value / Math.max(props.width, 1))
const chartVisuals = computed(() => new Map(props.slide.elements.flatMap(element => element.chart ? [[element.id, makeChartVisual(element.chart, element.id, element.width)] as const] : [])))
const chartPatternDefinitions = computed(() => new Map(props.slide.elements.flatMap(element => {
  if (!element.chart) return []
  const fills = [...new Set(element.chart.series.flatMap(series => [series.color, series.stroke, ...Object.values(series.pointColors), ...Object.values(series.pointStrokes || {})]).filter((fill): fill is string => fill !== undefined && isPatternFillCss(fill)))]
  return fills.length ? [[element.id, fills.map((fill, index) => ({ id: `${customGradientScope}-chart-pattern-${element.id.replace(/[^a-zA-Z0-9_-]/g, '_')}-${index}`, fill }))] as const] : []
})))
const chartPaint = (elementId: string, fill: string | undefined) => {
  if (!fill || !isPatternFillCss(fill)) return fill
  const pattern = chartPatternDefinitions.value.get(elementId)?.find(item => item.fill === fill)
  return pattern ? `url(#${pattern.id})` : fill
}
const chartDataLabelKeyPositions = ref(new Map<string, number>())
let chartDataLabelResizeObserver: ResizeObserver | undefined
const observedChartSvgs = new WeakSet<SVGSVGElement>()
const chartDataLabelKeyId = (elementId: string, index: number) => `${customGradientScope}:${elementId}:${index}`

const refreshChartDataLabelKeyPositions = async () => {
  await nextTick()
  if (typeof document === 'undefined') return
  const positions = new Map<string, number>()
  for (const key of document.querySelectorAll<SVGGElement>('.chart-data-label-key')) {
    const id = key.dataset.keyId
    if (!id?.startsWith(`${customGradientScope}:`)) continue
    const text = key.parentElement?.querySelector<SVGTextElement>('.chart-data-label')
    if (!text) continue
    const svg = key.closest('svg')
    if (svg && chartDataLabelResizeObserver && !observedChartSvgs.has(svg)) {
      chartDataLabelResizeObserver.observe(svg)
      observedChartSvgs.add(svg)
    }
    try {
      positions.set(id, text.textContent?.trim() ? text.getBBox().x - 18 : Number(text.getAttribute('x')) - 6)
    } catch {}
  }
  chartDataLabelKeyPositions.value = positions
}

watch(chartVisuals, () => { void refreshChartDataLabelKeyPositions() }, {flush: 'post'})
onMounted(() => {
  if (typeof ResizeObserver !== 'undefined') chartDataLabelResizeObserver = new ResizeObserver(() => { void refreshChartDataLabelKeyPositions() })
  void refreshChartDataLabelKeyPositions()
})
onBeforeUnmount(() => chartDataLabelResizeObserver?.disconnect())

function supportsTextVerticalEllipsis(element: PptxElement): boolean {
  return element.textVerticalOverflow === 'ellipsis'
    && (element.textColumnCount ?? 1) === 1
    && (!element.textOrientation || element.textOrientation === 'horz')
}

function measureTextEllipsisLineCount(frame: HTMLElement): number {
  const frameRect = frame.getBoundingClientRect()
  const paddingBottom = Number.parseFloat(getComputedStyle(frame).paddingBottom) || 0
  const contentBottom = frameRect.bottom - paddingBottom
  const bands: Array<{ top: number; bottom: number }> = []
  for (const run of frame.querySelectorAll<HTMLElement>('.text-run')) {
    const range = document.createRange()
    range.selectNodeContents(run)
    for (const rect of range.getClientRects()) {
      if (rect.width <= 0 || rect.height <= 0) continue
      const band = bands.find(line => rect.top < line.bottom - 1 && rect.bottom > line.top + 1)
      if (band) {
        band.top = Math.min(band.top, rect.top)
        band.bottom = Math.max(band.bottom, rect.bottom)
      } else {
        bands.push({ top: rect.top, bottom: rect.bottom })
      }
    }
  }
  bands.sort((left, right) => left.top - right.top)
  const visibleLines = bands.filter(line => line.bottom <= contentBottom + 0.5).length
  return bands.some(line => line.bottom > contentBottom + 0.5) ? Math.max(1, visibleLines) : 0
}

function scheduleShapeAutoFit(targetIds?: ReadonlySet<string>) {
  if (!stage.value || autoFitFrame || scale.value <= 0) return
  autoFitFrame = requestAnimationFrame(() => {
    autoFitFrame = 0
    if (!stage.value) return
    const isTargeted = (id: string | undefined) => Boolean(id && (!targetIds || targetIds.has(id)))
    const ellipsisFrames = [...stage.value.querySelectorAll<HTMLElement>('.text-frame[data-text-overflow="ellipsis"]')]
    let resetTextEllipsis = false
    let layoutChanged = false
    for (const frame of ellipsisFrames) {
      const id = frame.dataset.elementId
      if (!id) continue
      const deleted = textEllipsisLineCounts.delete(`${props.slide.id}:${id}`)
      layoutChanged ||= deleted
      resetTextEllipsis = deleted || resetTextEllipsis
    }
    if (resetTextEllipsis) {
      elementLayoutVersion.value++
      void nextTick(scheduleShapeAutoFit)
      return
    }
    const canClampText = typeof CSS !== 'undefined' && CSS.supports('display', '-webkit-box') && CSS.supports('-webkit-line-clamp', '1')
    textOverflowWarning.value = ellipsisFrames.length && !canClampText ? 'textEllipsisUnsupported' : ''
    if (canClampText) {
      for (const frame of ellipsisFrames) {
        const id = frame.dataset.elementId
        if (!isTargeted(id)) continue
        const lineCount = measureTextEllipsisLineCount(frame)
        if (lineCount > 0 && textEllipsisLineCounts.get(`${props.slide.id}:${id}`) !== lineCount) {
          textEllipsisLineCounts.set(`${props.slide.id}:${id}`, lineCount)
          layoutChanged = true
        }
      }
    }
    const elements = new Map(props.slide.elements.map(element => [element.id, element]))
    let changed = false
    for (const frame of stage.value.querySelectorAll<HTMLElement>('.text-frame[data-autofit="shape"]')) {
      const id = frame.dataset.elementId
      const element = id ? elements.get(id) : undefined
      if (!id || !isTargeted(id) || !element) continue
      const overflowWidth = Math.max(0, frame.scrollWidth - frame.clientWidth)
      const overflowHeight = Math.max(0, frame.scrollHeight - frame.clientHeight)
      if (!overflowWidth && !overflowHeight) continue
      const key = `${props.slide.id}:${id}`
      const current = autoFitDimensions.get(key) || { width: element.width, height: element.height }
      autoFitDimensions.set(key, {
        width: current.width + overflowWidth / scale.value,
        height: current.height + overflowHeight / scale.value,
      })
      changed = true
      layoutChanged = true
    }
    for (const frame of stage.value.querySelectorAll<HTMLElement>('.text-frame[data-autofit="normal"]')) {
      const id = frame.dataset.elementId
      const element = id ? elements.get(id) : undefined
      if (!id || !isTargeted(id) || !element || !frame.clientWidth || !frame.clientHeight) continue
      const overflowWidth = frame.scrollWidth > frame.clientWidth + 1 ? frame.clientWidth / frame.scrollWidth : 1
      const overflowHeight = frame.scrollHeight > frame.clientHeight + 1 ? frame.clientHeight / frame.scrollHeight : 1
      if (overflowWidth === 1 && overflowHeight === 1) continue
      const key = `${props.slide.id}:${id}`
      const current = normalAutoFitScales.get(key) ?? element.textFontScale ?? 1
      const fitted = Math.max(0.05, current * Math.min(overflowWidth, overflowHeight) * 0.98)
      if (fitted < current - 0.001) {
        normalAutoFitScales.set(key, fitted)
        changed = true
        layoutChanged = true
      }
    }
    for (const cell of stage.value.querySelectorAll<HTMLTableCellElement>('.slide-table td[data-autofit="normal"]')) {
      const elementNode = cell.closest<HTMLElement>('.slide-element')
      const id = elementNode?.dataset.elementId
      const rowIndex = Number(cell.dataset.rowIndex)
      const cellIndex = Number(cell.dataset.cellIndex)
      const element = id ? elements.get(id) : undefined
      const model = element?.table?.rows[rowIndex]?.cells[cellIndex]
      const text = cell.querySelector<HTMLElement>('.table-cell-text')
      if (!id || !isTargeted(id) || !model || !text?.clientWidth || !text.clientHeight) continue
      const overflowWidth = text.scrollWidth > text.clientWidth + 1 ? text.clientWidth / text.scrollWidth : 1
      const overflowHeight = text.scrollHeight > text.clientHeight + 1 ? text.clientHeight / text.scrollHeight : 1
      if (overflowWidth === 1 && overflowHeight === 1) continue
      const key = tableCellAutoFitKey(id, rowIndex, cellIndex)
      const current = normalAutoFitScales.get(key) ?? model.textFontScale ?? 1
      const fitted = Math.max(0.05, current * Math.min(overflowWidth, overflowHeight) * 0.98)
      if (fitted < current - 0.001) {
        normalAutoFitScales.set(key, fitted)
        changed = true
        layoutChanged = true
      }
    }
    if (layoutChanged) elementLayoutVersion.value++
    void nextTick(measureTextAnchorCenterOffsets)
    if (changed) void nextTick(scheduleShapeAutoFit)
  })
}

function measureTextAnchorCenterOffsets() {
  if (!stage.value) return
  const elements = new Map(props.slide.elements.map(element => [element.id, element]))
  let changed = false
  for (const frame of stage.value.querySelectorAll<HTMLElement>('.text-frame')) {
    const id = frame.dataset.elementId
    const element = id ? elements.get(id) : undefined
    if (!element?.textAnchorCenter || (element.textColumnCount ?? 1) <= 1) continue
    const bounds = frame.querySelector<HTMLElement>('.text-bounds')
    if (!bounds) continue

    const frameRect = frame.getBoundingClientRect()
    const frameStyle = getComputedStyle(frame)
    const paddingLeft = Number.parseFloat(frameStyle.paddingLeft) || 0
    const paddingRight = Number.parseFloat(frameStyle.paddingRight) || 0
    const paddingTop = Number.parseFloat(frameStyle.paddingTop) || 0
    const paddingBottom = Number.parseFloat(frameStyle.paddingBottom) || 0
    const parent = frame.closest<HTMLElement>('.slide-element')
    const transform = parent ? getComputedStyle(parent).transform : 'none'
    const matrix = transform === 'none' ? new DOMMatrixReadOnly() : new DOMMatrixReadOnly(transform)
    const targetX = frameRect.left + frameRect.width / 2 + matrix.a * (paddingLeft - paddingRight) / 2 + matrix.c * (paddingTop - paddingBottom) / 2
    const targetY = frameRect.top + frameRect.height / 2 + matrix.b * (paddingLeft - paddingRight) / 2 + matrix.d * (paddingTop - paddingBottom) / 2
    const key = `${props.slide.id}:${id}`
    const currentOffset = textAnchorCenterOffsets.get(key) ?? 0
    const boundsRect = bounds.getBoundingClientRect()
    const offsetX = targetX - (boundsRect.left + boundsRect.width / 2 - matrix.a * currentOffset)
    const offsetY = targetY - (boundsRect.top + boundsRect.height / 2 - matrix.b * currentOffset)
    const axisScale = matrix.a * matrix.a + matrix.b * matrix.b
    if (axisScale < 0.000001) continue
    const nextOffset = (offsetX * matrix.a + offsetY * matrix.b) / axisScale
    if (Math.abs(nextOffset) < 0.5) {
      changed = textAnchorCenterOffsets.delete(key) || changed
    } else if (Math.abs(nextOffset - currentOffset) >= 0.5) {
      textAnchorCenterOffsets.set(key, nextOffset)
      changed = true
    }
  }
  if (changed) elementLayoutVersion.value++
}

function onFontsLoaded() {
  autoFitDimensions.clear()
  normalAutoFitScales.clear()
  textEllipsisLineCounts.clear()
  elementLayoutVersion.value++
  scheduleShapeAutoFit()
}
const chartPlot = (id: string) => chartVisuals.value.get(id)?.plot ?? { left: 88, top: 75, width: 850, height: 390 }
type RenderAnimation = PptxAnimation & { step: number; automatic: boolean; triggerSequence?: number; triggerStep?: number }
const triggerKey = (sequence: number, step: number) => `${sequence}:${step}`
const timelineEnd = (action: RenderAnimation) => action.triggerSequence !== undefined && action.triggerStep !== undefined
  ? triggerStepEndedAt.get(triggerKey(action.triggerSequence, action.triggerStep))
  : action.automatic ? undefined : mainStepEndedAt.get(action.step)
const timelineBaseStart = (action: RenderAnimation) => action.triggerSequence === undefined || action.triggerStep === undefined
  ? action.automatic ? automaticStartedAt : stepStartedAt.get(action.step)
  : triggerStartedAt.get(triggerKey(action.triggerSequence, action.triggerStep))
const rawTimelineStart = (action: RenderAnimation) => {
  const startedAt = timelineBaseStart(action)
  if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY || !action.mediaWaitForEndKeys?.length) return startedAt
  const endedAt = action.mediaWaitForEndKeys.map(key => completedMediaActions.get(key))
  if (endedAt.some(time => time === undefined)) return Infinity
  return Math.max(startedAt, ...endedAt.map(time => time! + (action.mediaWaitDelayMs || 0) - action.delayMs))
}
const timelineNow = (action: RenderAnimation) => {
  const now = Math.min(animationClock.value, timelineEnd(action) ?? Infinity)
  return timelineTimeAt(action, now)
}
const animationDuration = (action: PptxAnimation) => action.durationMs === 0
  ? 0
  : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (action.repeatCount ?? 1), action.repeatDurationMs ?? Infinity)
const naturalAnimationDuration = (action: PptxAnimation) => action.effect === 'media'
  ? action.mediaDurationMs ?? 0
  : action.timingWarp
    ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs)
  : action.durationMs === 0
    ? 0
  : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (Number.isFinite(action.repeatCount) || action.repeatDurationMs !== undefined ? action.repeatCount ?? 1 : 1), action.repeatDurationMs ?? Infinity)
    + (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0)
const animationScheduleDuration = (action: PptxAnimation) => action.effect === 'media' ? 0 : action.timingWarp
  ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs)
  : animationDuration(action) + (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0)
function iterationIndexFor(action: PptxAnimation, start: number, end: number): number | undefined {
  const ranges = action.iteration?.ranges
  if (!ranges) return undefined
  let low = 0, high = ranges.length
  while (low < high) {
    const middle = (low + high) >>> 1
    if (ranges[middle]!.end <= start) low = middle + 1
    else high = middle
  }
  const range = ranges[low]
  return range && range.start <= start && range.end >= end
    ? action.iteration?.backwards ? ranges.length - low - 1 : low
    : -1
}
const animationsByTarget = computed(() => {
  const targets = new Map<string, RenderAnimation[]>()
  if (!props.showAnimations) return targets
  for (const animation of props.slide.automaticAnimations || []) {
    const actions = targets.get(animation.targetId) || []
    actions.push({ ...animation, step: -1, automatic: true })
    targets.set(animation.targetId, actions)
  }
  props.slide.animationSteps?.forEach((step, index) => {
    for (const animation of step) {
      const actions = targets.get(animation.targetId) || []
      actions.push({ ...animation, step: index, automatic: false })
      targets.set(animation.targetId, actions)
    }
  })
  props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
    sequence.steps.forEach((step, stepIndex) => {
      for (const animation of step) {
        const actions = targets.get(animation.targetId) || []
        actions.push({ ...animation, step: -1, automatic: false, triggerSequence: sequenceIndex, triggerStep: stepIndex })
        targets.set(animation.targetId, actions)
      }
    })
  })
  for (const actions of targets.values()) actions.sort((a, b) => Number(a.triggerSequence !== undefined) - Number(b.triggerSequence !== undefined) || a.step - b.step || a.delayMs - b.delayMs)
  return targets
})
const hasInteractiveMedia = computed(() => !props.thumbnail && props.slide.elements.some(element => element.kind === 'video' || element.kind === 'audio'))
const playbackActionEnabled = (action: PptxAnimation) => (props.showAnimations || action.effect === 'media') && (props.showNarration || !action.mediaIsNarration)
const timelineActions = computed<RenderAnimation[]>(() => [
  ...(props.slide.automaticAnimations || []).map(animation => ({ ...animation, step: -1, automatic: true })),
  ...(props.slide.animationSteps || []).flatMap((step, index) => step.map(animation => ({ ...animation, step: index, automatic: false }))),
  ...(props.slide.triggeredAnimations || []).flatMap((sequence, sequenceIndex) => sequence.steps.flatMap((step, stepIndex) => step.map(animation => ({ ...animation, step: -1, automatic: false, triggerSequence: sequenceIndex, triggerStep: stepIndex })))),
].filter(playbackActionEnabled))
const embeddedSoundActions = computed(() => props.thumbnail && !props.playbackPreview ? [] : timelineActions.value.filter(action => action.effect === 'media' && action.mediaSourceUrl))
const exclusiveTimelineRanges = computed(() => {
  // Keep Map-backed click starts reactive as the clock advances.
  const now = animationClock.value
  type ExclusiveGroup = {id: string; baseAt: number; startMs: number; endMs: number; parentId?: string}
  const groups = new Map<string, ExclusiveGroup>()
  const parentIds = new Map<string, string>()
  for (const action of timelineActions.value) {
    const startedAt = timelineBaseStart(action)
    if (startedAt === undefined || !Number.isFinite(startedAt)) continue
    const memberships = action.exclusiveGroups || []
    // The parser appends enclosing groups after their descendants, so membership order is inner-to-outer.
    for (let index = 0; index < memberships.length - 1; index++) {
      parentIds.set(memberships[index]!.id, memberships[index + 1]!.id)
    }
    for (const group of memberships) {
      if (!groups.has(group.id) && Number.isFinite(group.startMs) && Number.isFinite(group.endMs) && group.endMs > group.startMs) {
        groups.set(group.id, {id: group.id, baseAt: startedAt, startMs: group.startMs, endMs: group.endMs})
      }
    }
  }
  for (const group of groups.values()) group.parentId = parentIds.get(group.id)
  const childrenByParent = new Map<string, ExclusiveGroup[]>()
  for (const group of groups.values()) {
    if (!group.parentId || !groups.has(group.parentId)) continue
    const children = childrenByParent.get(group.parentId) || []
    children.push(group)
    childrenByParent.set(group.parentId, children)
  }
  const orderedChildren = (group: ExclusiveGroup) => (childrenByParent.get(group.id) || [])
    .sort((a, b) => a.baseAt + a.startMs - b.baseAt - b.startMs)
  const groupDurations = new Map<string, number>()
  const durationOf = (group: ExclusiveGroup, visiting = new Set<string>()): number => {
    const cached = groupDurations.get(group.id)
    if (cached !== undefined) return cached
    if (visiting.has(group.id)) return group.endMs - group.startMs
    const nextVisiting = new Set(visiting).add(group.id)
    // Nested exclusives pause the parent clock, extending its wall-clock range.
    const nestedPause = orderedChildren(group).reduce((total, child) => total + durationOf(child, nextVisiting), 0)
    const duration = group.endMs - group.startMs + nestedPause
    groupDurations.set(group.id, duration)
    return duration
  }
  const rootGroups = [...groups.values()]
    .filter(group => !group.parentId || !groups.has(group.parentId))
    .sort((a, b) => a.baseAt + a.startMs - b.baseAt - b.startMs)
  const ranges: Array<{id: string; startAt: number; endAt: number}> = []
  const scheduledRoots: Array<{id: string; startAt: number; endAt: number}> = []
  const scheduleChildren = (parent: ExclusiveGroup, parentStartAt: number) => {
    let accumulatedPauseMs = 0
    let previousLocalEnd = Number.NEGATIVE_INFINITY
    const parentNominalStart = parent.baseAt + parent.startMs
    for (const child of orderedChildren(parent)) {
      const childNominalOffset = child.baseAt + child.startMs - parentNominalStart
      let localStart = childNominalOffset + accumulatedPauseMs
      if (localStart < previousLocalEnd) {
        const siblingShift = previousLocalEnd - localStart
        localStart += siblingShift
        accumulatedPauseMs += siblingShift
      }
      const childStartAt = parentStartAt + localStart
      const childDuration = durationOf(child)
      const childRange = {id: child.id, startAt: childStartAt, endAt: childStartAt + childDuration}
      ranges.push(childRange)
      scheduleChildren(child, childStartAt)
      accumulatedPauseMs += childDuration
      previousLocalEnd = localStart + childDuration
    }
  }
  for (const group of rootGroups) {
    let startAt = group.baseAt + group.startMs
    const shifted = new Set<string>()
    while (true) {
      let adjusted = false
      for (const previous of scheduledRoots) {
        if (shifted.has(previous.id) || previous.startAt >= startAt || previous.endAt <= group.baseAt) continue
        startAt += previous.endAt - Math.max(previous.startAt, group.baseAt)
        shifted.add(previous.id)
        adjusted = true
      }
      if (!adjusted) break
    }
    const rootRange = {id: group.id, startAt, endAt: startAt + durationOf(group)}
    ranges.push(rootRange)
    scheduledRoots.push(rootRange)
    scheduleChildren(group, startAt)
  }
  return ranges.sort((a, b) => a.startAt - b.startAt || a.endAt - b.endAt)
})
function timelinePauseDuration(action: RenderAnimation, at: number): number {
  if (!Number.isFinite(at)) return 0
  const memberships = new Set(action.exclusiveGroups?.map(group => group.id) || [])
  const ranges = exclusiveTimelineRanges.value.filter(range => !memberships.has(range.id))
  let paused = 0
  let previousEnd = 0
  for (const range of ranges) {
    if (range.startAt >= at) break
    const start = Math.max(previousEnd, range.startAt, 0)
    const end = Math.min(at, range.endAt)
    if (end > start) paused += end - start
    previousEnd = Math.max(previousEnd, end)
  }
  return paused
}
function timelineTimeAt(action: RenderAnimation, time: number): number {
  return Number.isFinite(time) ? time - timelinePauseDuration(action, time) : time
}
const timelineStart = (action: RenderAnimation) => {
  const startedAt = rawTimelineStart(action)
  return startedAt === undefined || !Number.isFinite(startedAt) ? startedAt : timelineTimeAt(action, startedAt)
}
const timelineIsPaused = (action: RenderAnimation, time: number) => {
  const memberships = new Set(action.exclusiveGroups?.map(group => group.id) || [])
  return exclusiveTimelineRanges.value.some(range => !memberships.has(range.id) && range.startAt <= time && time < range.endAt)
}
function nextExclusiveTimelineBoundary(action: RenderAnimation, time: number): number {
  const memberships = new Set(action.exclusiveGroups?.map(group => group.id) || [])
  let boundary = Infinity
  for (const range of exclusiveTimelineRanges.value) {
    if (memberships.has(range.id)) continue
    if (range.startAt <= time && time < range.endAt) boundary = Math.min(boundary, range.endAt)
    else if (range.startAt > time) boundary = Math.min(boundary, range.startAt)
  }
  return boundary
}
function clearMediaActionEnd(actionKey: string) {
  const timer = mediaActionEndTimers.get(actionKey)
  if (timer) clearTimeout(timer)
  mediaActionEndTimers.delete(actionKey)
}
function scheduleMediaActionEnd(action: RenderAnimation, media: HTMLMediaElement, remainingMs: number, checkedAt = performance.now()) {
  const actionKey = action.mediaActionKey
  if (!actionKey) return
  const state = {action, media, remainingMs, checkedAt}
  const now = performance.now()
  const boundary = nextExclusiveTimelineBoundary(action, now)
  const delay = timelineIsPaused(action, now) ? boundary - now : Math.min(remainingMs, boundary - now)
  media.dataset.mediaActionEndAt = String(now + Math.max(0, remainingMs))
  mediaActionEndTimers.set(actionKey, setTimeout(() => {
    const endedAt = performance.now()
    const elapsed = endedAt - state.checkedAt
    const paused = timelinePauseDuration(action, endedAt) - timelinePauseDuration(action, state.checkedAt)
    state.remainingMs -= Math.max(0, elapsed - paused)
    if (state.remainingMs > 0) {
      scheduleMediaActionEnd(action, media, state.remainingMs, endedAt)
      return
    }
    clearMediaActionEnd(actionKey)
    delete media.dataset.mediaActionEndAt
    media.pause()
    finishTimedMedia(action.targetId)
  }, Math.max(0, delay)))
}
function syncExclusiveMediaPlayback(now: number) {
  const activeGroups = exclusiveTimelineRanges.value.filter(range => range.startAt <= now && now < range.endAt)
  if (!activeGroups.length && !exclusivePausedMedia.size) return
  const activeMediaActions = new Map<string, {action: RenderAnimation; startedAt: number}>()
  // A target can have pending actions in nested groups; only its latest started play action owns playback.
  for (const action of timelineActions.value) {
    const actionKey = action.mediaActionKey
    if (action.effect !== 'media' || !actionKey || !['play', 'togglePause'].includes(action.mediaCommand || '') || completedMediaActions.has(actionKey)) continue
    const startedAt = startedMediaActions.get(actionKey)
    if (startedAt === undefined) continue
    const previous = activeMediaActions.get(action.targetId)
    if (!previous || startedAt > previous.startedAt) activeMediaActions.set(action.targetId, {action, startedAt})
  }
  for (const media of stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || []) {
    if (media.ended) continue
    const mediaId = media.dataset.mediaId || ''
    const action = activeMediaActions.get(mediaId)?.action
    const memberships = new Set(action?.exclusiveGroups?.map(group => group.id) || [])
    // Nested media remains active only when it belongs to every currently exclusive ancestor/group.
    const belongsToEveryActiveGroup = activeGroups.length > 0 && activeGroups.every(group => memberships.has(group.id))
    if (activeGroups.length && !belongsToEveryActiveGroup) {
      if (!media.paused && !exclusivePausedMedia.has(media)) {
        const endAt = Number(media.dataset.mediaActionEndAt)
        exclusivePausedMedia.set(media, {
          remainingMs: Number.isFinite(endAt) ? endAt - now : undefined,
          actionKey: activeMediaActions.get(mediaId)?.action.mediaActionKey,
          pausedAt: now
        })
      }
      if (exclusivePausedMedia.has(media) && !media.paused) media.pause()
      continue
    }
    if (!exclusivePausedMedia.has(media)) continue
    const pausedPlayback = exclusivePausedMedia.get(media)
    exclusivePausedMedia.delete(media)
    if (!media.isConnected || !media.paused) continue
    const activePlayback = activeMediaActions.get(mediaId)
    const playbackRestartedWhilePaused = activePlayback
      && activePlayback.action.mediaActionKey !== pausedPlayback?.actionKey
      && activePlayback.startedAt >= (pausedPlayback?.pausedAt ?? Infinity)
    const pausedActionCompleted = pausedPlayback?.actionKey && completedMediaActions.has(pausedPlayback.actionKey)
    if (playbackRestartedWhilePaused || pausedActionCompleted) continue
    if (pausedPlayback?.remainingMs !== undefined) media.dataset.mediaActionEndAt = String(now + Math.max(0, pausedPlayback.remainingMs))
    void media.play().catch(error => {
      if (error instanceof DOMException && error.name === 'NotAllowedError' && mediaId && !blockedMediaPlaybackIds.value.includes(mediaId)) {
        blockedMediaPlaybackIds.value = [...blockedMediaPlaybackIds.value, mediaId]
      }
    })
  }
}
const triggerSequenceHasEnabledAction = (sequence: NonNullable<PptxSlide['triggeredAnimations']>[number]) => sequence.steps.some(step => step.some(playbackActionEnabled))
const triggerTargetIds = computed(() => {
  const targetIds = new Set<string>()
  if (props.thumbnail) return targetIds
  for (const sequence of props.slide.triggeredAnimations || []) {
    if (!triggerSequenceHasEnabledAction(sequence)) continue
    for (const condition of sequence.conditions) if (condition.type === 'shape') targetIds.add(condition.id)
  }
  return targetIds
})
function endTimelineStep(triggerSequence: number | undefined, step: number, endedAt: number) {
  if (triggerSequence === undefined) mainStepEndedAt.set(step, endedAt)
  else triggerStepEndedAt.set(triggerKey(triggerSequence, step), endedAt)
  timelineActions.value.forEach((action, index) => {
    const belongsToStep = triggerSequence === undefined
      ? action.triggerSequence === undefined && !action.automatic && action.step === step
      : action.triggerSequence === triggerSequence && action.triggerStep === step
    if (!belongsToStep || action.effect !== 'media' || !executedMediaActions.has(index)) return
    const media = Array.from(stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || [])
      .find(item => item.dataset.mediaId === action.targetId)
    if (media) {
      media.pause()
      delete media.dataset.mediaActionEndAt
      mediaBookmarkPlaybackStarted.delete(media)
      lastMediaBookmarkTimes.delete(media)
    }
    if (action.mediaActionKey) {
      completedMediaActions.set(action.mediaActionKey, endedAt)
      clearMediaActionEnd(action.mediaActionKey)
    }
  })
}
const mediaSlideCounts = computed(() => {
  const counts = new Map<string, number>()
  for (const action of timelineActions.value) {
    if (action.effect === 'media' && action.mediaSlideCount) counts.set(action.targetId, Math.max(action.mediaSlideCount, counts.get(action.targetId) || 0))
  }
  return counts
})

// shortcut: parsed slide targets stay immutable, invalidate on slide/visibility changes; add an animation revision if editing is introduced.
const elementMemoMetadata = new WeakMap<PptxElement, {
  slide: PptxSlide
  slideId: string
  showAnimations: boolean
  animations: RenderAnimation[] | undefined
  hasStepAnimations: boolean
  hasTriggerAnimations: boolean
  hasPotentialFillExpiration: boolean
  layoutKey: string
}>()

function elementMemoDependencies(element: PptxElement) {
  const slide = props.slide
  const slideId = slide.id
  let metadata = elementMemoMetadata.get(element)
  if (!metadata || metadata.slide !== slide || metadata.slideId !== slideId || metadata.showAnimations !== props.showAnimations) {
    const animations = animationsByTarget.value.get(element.id)
    const allActions = timelineActions.value
    const hasPotentialFillExpiration = Boolean(animations?.some(action => action.fillMode === 'freeze'
      && allActions.some(sibling => animationStartsLaterInFillScope(action, sibling))))
    metadata = {
      slide,
      slideId,
      showAnimations: props.showAnimations,
      animations,
      hasStepAnimations: Boolean(animations?.some(action => !action.automatic && action.triggerSequence === undefined)),
      hasTriggerAnimations: Boolean(animations?.some(action => action.triggerSequence !== undefined)),
      hasPotentialFillExpiration,
      layoutKey: `${slideId}:${element.id}`,
    }
    elementMemoMetadata.set(element, metadata)
  }
  const {animations, hasStepAnimations, hasTriggerAnimations, hasPotentialFillExpiration, layoutKey} = metadata
  const step = hasStepAnimations ? props.animationStep : null
  const animationTime = animations?.length && (hasTriggerAnimations || hasPotentialFillExpiration || animationClockTargetIds.has(element.id))
    ? animationClock.value
    : null
  const timelineVersion = animations?.length ? animationTimelineVersion.value : null
  const fitted = autoFitDimensions.get(layoutKey)
  const fontScale = normalAutoFitScales.get(layoutKey)
  const ellipsisLines = textEllipsisLineCounts.get(layoutKey)
  const mediaSlideCount = mediaSlideCounts.value.get(element.id)
  const invalidTrim = invalidMediaTrimIds.value.includes(element.id)
  const invalidFade = invalidMediaFadeIds.value.includes(element.id)
  const blockedPlayback = blockedMediaPlaybackIds.value.includes(element.id)
  return [
    props.slide,
    slideId,
    element,
    props.width,
    props.height,
    step,
    props.thumbnail,
    props.playbackPreview,
    props.triggerPlayback,
    scale.value,
    pixelRatio.value,
    elementLayoutVersion.value,
    fitted,
    fontScale,
    ellipsisLines,
    mediaSlideCount,
    invalidTrim,
    invalidFade,
    blockedPlayback,
    animations,
    animationTime,
    timelineVersion,
  ]
}

function cancelTimelineUpdates() {
  cancelAnimationFrame(animationFrame)
  animationFrame = 0
  if (animationTimer) clearTimeout(animationTimer)
  animationTimer = undefined
}

function scheduleTimelineUpdate() {
  cancelTimelineUpdates()
  animationClockTargetIds.clear()
  if (props.thumbnail && !props.playbackPreview) return
  runTimedMediaActions()
  const now = animationClock.value
  syncExclusiveMediaPlayback(now)
  const exclusiveActive = exclusiveTimelineRanges.value.some(range => range.startAt <= now && now < range.endAt)
  const nextTimingNodeTriggerAt = exclusiveActive ? Infinity : runTimingNodeTriggeredAnimations(now)
  const scheduleAnimatedTextAutoFit = () => {
    const targetIds = new Set(timelineActions.value
      .filter(action => action.effect === 'fontSize' && animationClockTargetIds.has(action.targetId))
      .map(action => action.targetId))
    if (targetIds.size) scheduleShapeAutoFit(targetIds)
  }
  if (reduceMotion.value) {
    for (const targetId of animationsByTarget.value.keys()) animationClockTargetIds.add(targetId)
    const nextScheduledAt = timelineActions.value.reduce((next, action) => {
      if (action.effect !== 'media') return next
      const startedAt = timelineStart(action)
      if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY) return next
      return Math.min(next, now + Math.max(0, startedAt + action.delayMs - timelineNow(action)))
    }, nextTimingNodeTriggerAt)
    const nextExclusiveBoundary = exclusiveTimelineRanges.value.reduce((next, range) => {
      const boundary = range.startAt > now ? range.startAt : range.endAt > now ? range.endAt : Infinity
      return Math.min(next, boundary)
    }, Infinity)
    scheduleAnimatedTextAutoFit()
    const nextScheduledTime = Math.min(nextScheduledAt, nextExclusiveBoundary)
    if (Number.isFinite(nextScheduledTime) && nextScheduledTime > now) {
      animationTimer = setTimeout(() => {
        animationTimer = undefined
        animationClock.value = performance.now()
        scheduleTimelineUpdate()
      }, Math.max(0, nextScheduledTime - now))
    }
    return
  }
  let nextStart = nextTimingNodeTriggerAt
  let hasActiveAnimation = exclusiveActive
  for (const action of timelineActions.value) {
    const startedAt = timelineStart(action)
    if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY) continue
    const stoppedAt = timelineEnd(action)
    const actionNow = timelineNow(action)
    const actionStop = stoppedAt === undefined ? undefined : timelineTimeAt(action, stoppedAt)
    if (actionStop !== undefined && actionStop <= actionNow) continue
    const start = startedAt + action.delayMs
    const end = Math.min(start + animationScheduleDuration(action), actionStop ?? Infinity)
    if (start > actionNow) {
      nextStart = Math.min(nextStart, now + start - actionNow)
      animationClockTargetIds.add(action.targetId)
    } else {
      if (end > actionNow) hasActiveAnimation = true
      if (end > actionNow || start === actionNow) animationClockTargetIds.add(action.targetId)
    }
  }
  scheduleAnimatedTextAutoFit()
  const update = () => {
    animationFrame = 0
    animationTimer = undefined
    animationClock.value = performance.now()
    scheduleTimelineUpdate()
  }
  if (hasActiveAnimation) animationFrame = requestAnimationFrame(update)
  else if (Number.isFinite(nextStart)) animationTimer = setTimeout(update, Math.max(0, nextStart - now))
}

function refreshTimeline() {
  cancelTimelineUpdates()
  animationTimelineVersion.value++
  animationClock.value = performance.now()
  scheduleTimelineUpdate()
}

function beginSlideTimeline(step: number) {
  const now = performance.now()
  invalidMediaTrimIds.value = []
  invalidMediaFadeIds.value = []
  blockedMediaPlaybackIds.value = []
  executedMediaActions.clear()
  scheduledMediaActions.clear()
  startedMediaActions.clear()
  completedMediaActions.clear()
  endedMediaTargets.clear()
  for (const actionKey of mediaActionEndTimers.keys()) clearMediaActionEnd(actionKey)
  automaticStartedAt = now
  stepStartedAt = new Map(Array.from({ length: step }, (_, index) => [index, Number.NEGATIVE_INFINITY]))
  mainStepEndedAt.clear()
  triggerSteps.clear()
  triggerStartedAt.clear()
  triggerStepEndedAt.clear()
  triggerPlaybackByKey.clear()
  firedTimingNodeTriggers.clear()
  triggerConditionDueAt.clear()
  pendingTriggerConditions.clear()
  animationTimelineVersion.value++
  animationClock.value = now
  scheduleTimelineUpdate()
}

function syncTriggerPlayback() {
  if (props.triggerPlayback === undefined) return
  triggerSteps.clear()
  triggerStartedAt.clear()
  triggerStepEndedAt.clear()
  triggerPlaybackByKey.clear()
  for (const playback of [...props.triggerPlayback].sort((a, b) => a.startedAt - b.startedAt)) {
    const sequence = props.slide.triggeredAnimations?.[playback.sequence]
    if (!sequence || !Number.isInteger(playback.step) || playback.step < 0 || playback.step >= sequence.steps.length || !Number.isFinite(playback.startedAt)) continue
    const key = triggerKey(playback.sequence, playback.step)
    const startedAt = playback.startedAt - performance.timeOrigin
    if (!sequence.concurrent && playback.step > 0) {
      const previousKey = triggerKey(playback.sequence, playback.step - 1)
      if (triggerStartedAt.has(previousKey)) endTimelineStep(playback.sequence, playback.step - 1, startedAt)
    }
    triggerStartedAt.set(key, startedAt)
    triggerPlaybackByKey.set(key, playback)
    triggerSteps.set(playback.sequence, Math.max(triggerSteps.get(playback.sequence) || 0, playback.step + 1))
  }
  refreshTimeline()
}

function updatePixelRatio() {
  pixelRatioQuery?.removeEventListener('change', updatePixelRatio)
  pixelRatio.value = window.devicePixelRatio || 1
  pixelRatioQuery = window.matchMedia(`(resolution: ${pixelRatio.value}dppx)`)
  pixelRatioQuery.addEventListener('change', updatePixelRatio)
}

onMounted(() => {
  if (!stage.value) return
  updatePixelRatio()
  window.addEventListener('resize', updatePixelRatio)
  motionPreference = window.matchMedia('(prefers-reduced-motion: reduce)')
  reduceMotion.value = motionPreference.matches
  motionPreference.addEventListener('change', onMotionPreferenceChange)
  const update = () => {
    stageWidth.value = stage.value?.getBoundingClientRect().width || 0
    void nextTick(scheduleShapeAutoFit)
  }
  observer = new ResizeObserver(update)
  observer.observe(stage.value)
  document.fonts.addEventListener('loadingdone', onFontsLoaded)
  void document.fonts.ready.then(onFontsLoaded)
  update()
  scheduleShapeAutoFit()
  beginSlideTimeline(props.animationStep)
})

watch([() => props.slide.id, () => props.animationStep], ([slideId, step], [oldSlideId, oldStep]) => {
  if (!stage.value) return
  if (slideId !== oldSlideId) {
    autoFitDimensions.clear()
    normalAutoFitScales.clear()
    textEllipsisLineCounts.clear()
    textAnchorCenterOffsets.clear()
    beginSlideTimeline(step)
    void nextTick(scheduleShapeAutoFit)
    return
  }
  if (step > oldStep) {
    for (let previous = oldStep + 1; previous < step; previous++) stepStartedAt.set(previous, Number.NEGATIVE_INFINITY)
    if (step === oldStep + 1) runSlideTriggeredAnimations('onNext')
    const now = performance.now()
    const previousStep = oldStep - 1
    const stepSequence = previousStep >= 0 ? props.slide.animationStepSequences?.[previousStep] : undefined
    const concurrent = stepSequence?.concurrent ?? props.slide.animationSequence?.concurrent
    const nextAction = stepSequence?.nextAction ?? props.slide.animationSequence?.nextAction
    if (previousStep >= 0 && !concurrent && nextAction !== 'seek') {
      endTimelineStep(undefined, previousStep, now)
    }
    if (nextAction === 'seek' && oldStep > 0) {
      const startedAt = stepStartedAt.get(previousStep)
      const naturalEnd = Math.max(0, ...(props.slide.animationSteps?.[previousStep] || []).map(action => action.delayMs + naturalAnimationDuration(action)))
      if (startedAt !== undefined && startedAt !== Number.NEGATIVE_INFINITY) stepStartedAt.set(previousStep, Math.min(startedAt, now - naturalEnd))
    }
    stepStartedAt.set(oldStep, now)
  } else if (step < oldStep) {
    if (step === oldStep - 1) runSlideTriggeredAnimations('onPrev')
    props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
      sequence.conditions.forEach((condition, conditionIndex) => {
        if (condition.type !== 'timeNode' || condition.sourceStep === undefined || condition.sourceStep < step) return
        const key = triggerConditionKey(sequenceIndex, conditionIndex)
        firedTimingNodeTriggers.delete(key)
        triggerConditionDueAt.delete(key)
        pendingTriggerConditions.delete(key)
      })
    })
    for (const startedStep of stepStartedAt.keys()) {
      if (startedStep >= step) {
        stepStartedAt.delete(startedStep)
        mainStepEndedAt.delete(startedStep)
      }
    }
    timelineActions.value.forEach((action, index) => {
      if (action.triggerSequence === undefined && !action.automatic && action.step >= step) executedMediaActions.delete(index)
    })
  }
  refreshTimeline()
}, { flush: 'post' })
watch([() => props.slide.id, () => props.triggerPlayback], syncTriggerPlayback, { flush: 'post' })
onUnmounted(() => {
  observer?.disconnect()
  pixelRatioQuery?.removeEventListener('change', updatePixelRatio)
  window.removeEventListener('resize', updatePixelRatio)
  document.fonts.removeEventListener('loadingdone', onFontsLoaded)
  cancelAnimationFrame(autoFitFrame)
  textEllipsisLineCounts.clear()
  textAnchorCenterOffsets.clear()
  motionPreference?.removeEventListener('change', onMotionPreferenceChange)
  cancelTimelineUpdates()
  for (const frame of mediaPlaybackFrames.values()) cancelAnimationFrame(frame)
  mediaPlaybackFrames.clear()
  for (const actionKey of mediaActionEndTimers.keys()) clearMediaActionEnd(actionKey)
  for (const audio of stage.value?.querySelectorAll<HTMLAudioElement>('audio[data-animation-sound]') || []) {
    audio.pause()
    if (audio.readyState > 0) audio.currentTime = 0
  }
  exclusivePausedMedia.clear()
})

function onMotionPreferenceChange(event: MediaQueryListEvent) {
  reduceMotion.value = event.matches
  refreshTimeline()
}

// Microsoft lists fill.opacity as an animatable property; multiplying each gradient stop alpha is this renderer's mapping and still needs Office/WPS screenshot calibration.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.attributename?view=openxml-3.0.1
// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/981b17ff-5594-42cf-ad8d-7cb39e653afa
function applyColorOpacity(color: string, opacity: number, brightness = 0): string {
  if ((opacity >= 1 && !brightness) || color === 'transparent') return color
  if (/(?:repeating-)?(?:linear|radial|conic)-gradient\(/i.test(color)) {
    return color.replace(/#[\da-f]{6}\b|rgba?\([^)]*\)|\btransparent\b/gi, stopColor => applyColorOpacity(stopColor, opacity, brightness))
  }
  const hex = /^#([\da-f]{6})$/i.exec(color)
  const rgb = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)(?:\s*,\s*([\d.]+)\s*)?\)$/i.exec(color)
  const channels = hex
    ? [parseInt(hex[1]!.slice(0, 2), 16), parseInt(hex[1]!.slice(2, 4), 16), parseInt(hex[1]!.slice(4, 6), 16), 1]
    : rgb ? [Number(rgb[1]), Number(rgb[2]), Number(rgb[3]), rgb[4] === undefined ? 1 : Number(rgb[4])] : undefined
  if (!channels) return color
  const rgbChannels = channels.slice(0, 3).map(channel => brightness ? Math.round(channel * (1 - Math.abs(brightness)) + Math.max(0, brightness) * 255) : channel)
  return `rgba(${rgbChannels.join(', ')}, ${Math.max(0, Math.min(channels[3]! * opacity, 1))})`
}

function animatedRunColor(color: string | undefined, opacity: number): string | undefined {
  if (opacity >= 1) return color
  const value = color || 'currentColor'
  const adjusted = applyColorOpacity(value, opacity)
  return adjusted !== value || value === 'transparent'
    ? adjusted
    : `color-mix(in srgb, ${value} ${Math.max(0, Math.min(opacity, 1)) * 100}%, transparent)`
}

function elementRenderSize(element: PptxElement): { width: number; height: number; scaleX: number; scaleY: number } {
  const fittedSize = autoFitDimensions.get(`${props.slide.id}:${element.id}`)
  const matrix = element.renderMatrix
  const matrixScaleX = matrix ? Math.hypot(matrix[0], matrix[1]) : 1
  const matrixScaleY = matrix ? Math.hypot(matrix[2], matrix[3]) : 1
  const scaleX = Number.isFinite(matrixScaleX) && matrixScaleX > 0 ? matrixScaleX : 1
  const scaleY = Number.isFinite(matrixScaleY) && matrixScaleY > 0 ? matrixScaleY : 1
  return {
    width: (fittedSize?.width || element.width) * scaleX,
    height: (fittedSize?.height || element.height) * scaleY,
    scaleX,
    scaleY,
  }
}

function elementStyle(element: PptxElement): CSSProperties {
  const { width, height, scaleX: matrixScaleX, scaleY: matrixScaleY } = elementRenderSize(element)
  const flips = [element.flipH ? 'scaleX(-1)' : '', element.flipV ? 'scaleY(-1)' : ''].filter(Boolean).join(' ')
  const matrix = element.renderMatrix
  const [scaleX, scaleY] = animationScale(element)
  const [motionX, motionY] = animationMotion(element)
  const [slideX, slideY] = animationSlide(element)
  const hasScaleAnimation = scaleX !== 1 || scaleY !== 1
  const rotation = animationRotation(element)
  const rotationDelta = rotation - element.rotation
  const hasRotationAnimation = Math.abs(rotationDelta) > 0.0001
  const motion = motionX || motionY ? `translate(${motionX * scale.value}px,${motionY * scale.value}px)` : ''
  const slideMotion = slideX || slideY ? `translate(${slideX * scale.value}px,${slideY * scale.value}px)` : ''
  const translations = [motion, slideMotion].filter(Boolean).join(' ')
  const shadows = element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []
  const fillOn = animationOnOff(element, 'fill.on')
  const strokeOn = animationOnOff(element, 'stroke.on')
  const shadowOn = animationOnOff(element, 'shadow.on')
  const shadowBlurScale = element.kind === 'shape' && element.geometry === 'ellipse' ? 0.5 : 1
  const renderBlurScale = Math.sqrt(matrixScaleX * matrixScaleY)
  const shadowColors = shadows.map((shadow, index) => shadowCss(shadow, applyColorOpacity(animationColor(element, 'shadow.color') || shadow.color, animationOpacity(element, 'shadow.opacity')), shadow.inset ? 1 : shadowBlurScale, { scaleX: matrixScaleX, scaleY: matrixScaleY }, animationShadowOffset(element, shadow, index)))
  const animatedFill = animationColor(element, 'fillcolor')
  const fill = animatedFill || element.fill
  const fillOpacity = animationOpacity(element, 'fill.opacity')
  const animatedStroke = animationColor(element, 'stroke.color')
  const stroke = applyColorOpacity(animatedStroke || element.stroke, animationOpacity(element, 'stroke.opacity'))
  const strokeWidth = animationStrokeWidth(element)
  const strokeDash = animationStrokeDash(element)
  const glows = element.glows?.length ? element.glows : element.glow ? [element.glow] : []
  const customGeometry = Boolean(element.customPaths?.length)
  const radialFill = fillOn && !animatedFill && !customGeometry && !usesSvgPolygon(element) ? customSvgGradient(element, 'fill') : undefined
  const radialBackground = radialFill?.kind === 'radial' && radialFill.fx !== undefined && radialFill.fy !== undefined && radialFill.fr !== undefined
    ? radialGradientBackground(radialFill)
    : undefined
  const showStroke = strokeOn && element.kind === 'shape' && !customGeometry && !usesSvgOutline(element) && strokeWidth > 0 && stroke !== 'transparent' && !stroke.includes('gradient')
  const normalizedMatrix = matrix
    ? [matrix[0] / matrixScaleX, matrix[1] / matrixScaleX, matrix[2] / matrixScaleY, matrix[3] / matrixScaleY]
    : undefined
  const transform = matrix
    ? [translations, `matrix(${normalizedMatrix!.join(',')},${matrix[4] * scale.value},${matrix[5] * scale.value})${hasScaleAnimation || hasRotationAnimation ? ` translate(${width * scale.value / 2}px,${height * scale.value / 2}px)${hasRotationAnimation ? ` rotate(${rotationDelta}deg)` : ''}${hasScaleAnimation ? ` scale(${scaleX},${scaleY})` : ''} translate(${-width * scale.value / 2}px,${-height * scale.value / 2}px)` : ''}`].filter(Boolean).join(' ')
    : [translations, rotation ? `rotate(${rotation}deg)` : '', flips, hasScaleAnimation ? `scale(${scaleX},${scaleY})` : ''].filter(Boolean).join(' ')
  return {
    left: `${(element.x / props.width) * 100}%`,
    top: `${(element.y / props.height) * 100}%`,
    width: `${(width / props.width) * 100}%`,
    height: `${(height / props.height) * 100}%`,
    overflow: customGeometry || usesSvgOutline(element) ? 'visible' : undefined,
    transform: transform || undefined,
    transformOrigin: matrix ? '0 0' : undefined,
    background: element.kind === 'shape' && !customGeometry && !usesSvgPolygon(element) ? fillOn ? radialBackground || applyColorOpacity(fill, fillOpacity) : 'none' : undefined,
    borderColor: showStroke ? stroke : undefined,
    borderWidth: showStroke ? `${strokeWidth * scale.value * renderBlurScale}px` : undefined,
    borderStyle: showStroke ? cssLineStyle(strokeDash) : undefined,
    boxShadow: shadows.length && shadowOn && !customGeometry && !usesSvgPolygon(element) ? shadowColors.join(', ') : undefined,
    filter: glows.length ? glows.map(glow => `drop-shadow(0 0 ${glow.radius * scale.value * renderBlurScale}px ${glow.color})`).join(' ') : undefined,
  }
}

function usesSvgOutline(element: PptxElement): boolean {
  return element.kind === 'shape' && (['rect', 'ellipse', 'roundRect', 'flowChartConnector', 'flowChartSummingJunction', 'flowChartOr', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDelay', 'flowChartDisplay', 'flowChartInternalStorage'].includes(element.geometry) || usesSvgPolygon(element))
}

function usesSvgPolygon(element: PptxElement): boolean {
  return element.kind === 'shape' && isSvgPolygonGeometry(element.geometry)
}

function presetBracketBracePaths(element: PptxElement): [string, string] {
  const rectangle = 'M0 0 H1000 V1000 H0 Z'
  if (!(element.width > 0) || !(element.height > 0)) return [rectangle, rectangle]
  const shortSide = Math.min(element.width, element.height)
  if (element.geometry === 'leftBrace' || element.geometry === 'rightBrace') {
    const adjustment2 = Math.max(0, Math.min(100_000, element.presetAdjustments?.adj2 ?? 50_000))
    const maximumAdjustment1 = Math.min(adjustment2, 100_000 - adjustment2) / 2 * element.height / shortSide
    const adjustment1 = Math.max(0, Math.min(maximumAdjustment1, element.presetAdjustments?.adj1 ?? 8_333))
    const radiusY = shortSide * adjustment1 / 100_000
    const middleX = element.width / 2
    const y3 = element.height * adjustment2 / 100_000
    const y2 = y3 - radiusY, y4 = element.height - radiusY
    const buildPath = (closed: boolean) => {
      let current: [number, number] = [0, 0]
      const point = ([x, y]: [number, number]) => `${Number((x / element.width * 1000).toFixed(6))} ${Number((y / element.height * 1000).toFixed(6))}`
      const move = (x: number, y: number) => { current = [x, y]; return `M${point(current)}` }
      const line = (x: number, y: number) => { current = [x, y]; return `L${point(current)}` }
      // ECMA-376 brace guides use DrawingML arcs locked to the current point; positive sweeps are clockwise.
      const arc = (radiusX: number, startAngle: number, sweepAngle: number) => {
        if (!(radiusX > 0) || !(radiusY > 0)) return ''
        const start = startAngle * Math.PI / 180, end = (startAngle + sweepAngle) * Math.PI / 180
        const centerX = current[0] - radiusX * Math.cos(start), centerY = current[1] - radiusY * Math.sin(start)
        current = [centerX + radiusX * Math.cos(end), centerY + radiusY * Math.sin(end)]
        return `A${Number((radiusX / element.width * 1000).toFixed(6))} ${Number((radiusY / element.height * 1000).toFixed(6))} 0 ${Math.abs(sweepAngle) > 180 ? 1 : 0} ${sweepAngle >= 0 ? 1 : 0} ${point(current)}`
      }
      const commands = element.geometry === 'leftBrace'
        ? [move(element.width, element.height), arc(middleX, 90, 90), line(middleX, y3 - radiusY), arc(middleX, 0, -90), arc(middleX, 90, -90), line(middleX, radiusY), arc(middleX, 180, 90)]
        : [move(0, 0), arc(middleX, 270, 90), line(middleX, y2), arc(middleX, 180, -90), arc(middleX, 270, -90), line(middleX, y4), arc(middleX, 0, 90)]
      if (closed) commands.push('Z')
      return commands.filter(Boolean).join(' ')
    }
    return [buildPath(true), buildPath(false)]
  }
  const isPair = element.geometry === 'bracketPair'
  if (element.geometry === 'bracePair') {
    const adjustment = Math.max(0, Math.min(25_000, element.presetAdjustments?.adj ?? 8_333))
    const radius = shortSide * adjustment / 100_000
    const x2 = shortSide * adjustment / 50_000
    const x3 = element.width - x2, x4 = element.width - radius
    const y2 = element.height / 2 - radius, y3 = element.height / 2 + radius, y4 = element.height - radius
    const point = ([x, y]: [number, number]) => `${Number((x / element.width * 1000).toFixed(6))} ${Number((y / element.height * 1000).toFixed(6))}`
    const buildPath = (outline: boolean) => {
      let current: [number, number] = [0, 0]
      const move = (x: number, y: number) => { current = [x, y]; return `M${point(current)}` }
      const line = (x: number, y: number) => { current = [x, y]; return `L${point(current)}` }
      // ECMA-376 bracePair uses DrawingML arcs locked to the current point; positive sweeps are clockwise.
      const arc = (startAngle: number, sweepAngle: number) => {
        if (!(radius > 0)) return ''
        const start = startAngle * Math.PI / 180, end = (startAngle + sweepAngle) * Math.PI / 180
        const centerX = current[0] - radius * Math.cos(start), centerY = current[1] - radius * Math.sin(start)
        current = [centerX + radius * Math.cos(end), centerY + radius * Math.sin(end)]
        return `A${Number((radius / element.width * 1000).toFixed(6))} ${Number((radius / element.height * 1000).toFixed(6))} 0 ${Math.abs(sweepAngle) > 180 ? 1 : 0} ${sweepAngle >= 0 ? 1 : 0} ${point(current)}`
      }
      const leftBrace = [move(x2, element.height), arc(90, 90), line(radius, y3), arc(0, -90), arc(90, -90), line(radius, radius), arc(180, 90)]
      const rightBrace = [move(x3, 0), arc(270, 90), line(x4, y2), arc(180, -90), arc(270, -90), line(x4, y4), arc(0, 90)]
      const joinAtTop = `L${point([x3, 0])}`
      return (outline ? [...leftBrace, ...rightBrace] : [...leftBrace, joinAtTop, ...rightBrace.slice(1), 'Z']).filter(Boolean).join(' ')
    }
    return [buildPath(false), buildPath(true)]
  }
  const maximumAdjustment = isPair ? 50_000 : 50_000 * element.height / shortSide
  const adjustment = Math.max(0, Math.min(maximumAdjustment, element.presetAdjustments?.adj ?? (isPair ? 16_667 : 8_333)))
  if (isPair) {
    const inset = shortSide * adjustment / 100_000 * 29_289 / 100_000
    const x = inset / element.width * 1000, y = inset / element.height * 1000
    const oppositeX = 1000 - x, oppositeY = 1000 - y
    const fill = `M0 ${y} A${x} ${y} 0 0 1 ${x} 0 L${oppositeX} 0 A${x} ${y} 0 0 1 1000 ${y} L1000 ${oppositeY} A${x} ${y} 0 0 1 ${oppositeX} 1000 L${x} 1000 A${x} ${y} 0 0 1 0 ${oppositeY} Z`
    const outline = `M${x} 1000 A${x} ${y} 0 0 1 0 ${oppositeY} L0 ${y} A${x} ${y} 0 0 1 ${x} 0 M${oppositeX} 0 A${x} ${y} 0 0 1 1000 ${y} L1000 ${oppositeY} A${x} ${y} 0 0 1 ${oppositeX} 1000`
    return [fill, outline]
  }
  const curveHeight = shortSide * adjustment / 100_000 / element.height * 1000
  const lowerCurveStart = 1000 - curveHeight
  const outline = element.geometry === 'rightBracket'
    ? `M0 0 A1000 ${curveHeight} 0 0 1 1000 ${curveHeight} L1000 ${lowerCurveStart} A1000 ${curveHeight} 0 0 1 0 1000`
    : `M1000 1000 A1000 ${curveHeight} 0 0 1 0 ${lowerCurveStart} L0 ${curveHeight} A1000 ${curveHeight} 0 0 1 1000 0`
  return [`${outline} Z`, outline]
}

function presetOutlineDetailPaths(element: PptxElement): string[] {
  if (element.geometry === 'leftBracket' || element.geometry === 'rightBracket' || element.geometry === 'bracketPair' || element.geometry === 'leftBrace' || element.geometry === 'rightBrace' || element.geometry === 'bracePair') return [presetBracketBracePaths(element)[1]]
  if (element.geometry === 'arc') {
    const arc = presetAngleArcPaths(element)[1]
    return arc ? [arc] : []
  }
  if (element.geometry === 'frame') return [presetOutlinePath(element)]
  if (element.geometry === 'cube') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return [presetOutlinePath(element)]
    const inset = Math.min(width, height) * (element.presetAdjustment ?? 0.25)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return [presetOutlinePath(element), `M0 ${y(inset)} L${x(width - inset)} ${y(inset)} L1000 0`, `M${x(width - inset)} ${y(inset)} L${x(width - inset)} 1000`]
  }
  if (element.geometry === 'flowChartMagneticDisk') return [presetOutlinePath(element), 'M1000 166.6667 A500 166.6667 0 0 1 0 166.6667']
  if (element.geometry === 'flowChartMagneticDrum') {
    const inset = 1000 / 6
    return [presetOutlinePath(element), `M${1000 - inset} 1000 A${inset} 500 0 0 1 ${1000 - inset} 0`]
  }
  if (element.geometry === 'flowChartMultidocument') {
    const scale = 1000 / 21600
    const point = (x: number, y: number) => `${x * scale} ${y * scale}`
    return [
      `M${point(0, 3675)} L${point(18595, 3675)} L${point(18595, 18022)} C${point(9298, 18022)} ${point(9298, 23542)} ${point(0, 20782)} Z`,
      `M${point(1532, 3675)} L${point(1532, 1815)} L${point(20000, 1815)} L${point(20000, 16252)} C${point(19298, 16252)} ${point(18595, 16352)} ${point(18595, 16352)}`,
      `M${point(2972, 1815)} L${point(2972, 0)} L${point(21600, 0)} L${point(21600, 14392)} C${point(20800, 14392)} ${point(20000, 14467)} ${point(20000, 14467)}`,
    ]
  }
  return []
}

function roundRectRadii(element: PptxElement): [number, number] {
  const fitted = autoFitDimensions.get(`${props.slide.id}:${element.id}`)
  const width = fitted?.width || element.width, height = fitted?.height || element.height
  const radius = Math.min(width, height) * (element.presetAdjustment ?? 0.16667)
  return [width > 0 ? radius / width * 1000 : 0, height > 0 ? radius / height * 1000 : 0]
}

function presetAngleArcPaths(element: PptxElement): [string, string] {
  const width = element.width, height = element.height
  if (width <= 0 || height <= 0) return ['M0 0 H1000 V1000 H0 Z', '']
  const fullTurn = 21_600_000
  const startAngle = element.presetAdjustments?.adj1 ?? (element.geometry === 'pie' ? 0 : element.geometry === 'chord' ? 2_700_000 : 16_200_000)
  const endAngle = element.presetAdjustments?.adj2 ?? (element.geometry === 'arc' ? 0 : 16_200_000)
  let sweepAngle = endAngle - startAngle
  if (sweepAngle <= 0) sweepAngle += fullTurn
  const [startX, startY] = presetEllipsePoint(width, height, startAngle)
  const arc = sweepAngle === fullTurn
    ? (() => {
        const [middleX, middleY] = presetEllipsePoint(width, height, startAngle + fullTurn / 2)
        return `M${startX} ${startY} A500 500 0 0 1 ${middleX} ${middleY} A500 500 0 0 1 ${startX} ${startY}`
      })()
    : (() => {
        const [endX, endY] = presetEllipsePoint(width, height, endAngle)
        return `M${startX} ${startY} A500 500 0 ${sweepAngle > fullTurn / 2 ? 1 : 0} 1 ${endX} ${endY}`
      })()
  return [element.geometry === 'chord' ? `${arc} Z` : `${arc} L500 500 Z`, arc]
}

function presetBlockArcPath(element: PptxElement): string {
  const width = element.width, height = element.height
  if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
  const fullTurn = 21_600_000
  const startAngle = element.presetAdjustments?.adj1 ?? 10_800_000
  const endAngle = element.presetAdjustments?.adj2 ?? 0
  let sweepAngle = endAngle - startAngle
  if (sweepAngle <= 0) sweepAngle += fullTurn
  const inset = Math.min(width, height) * (element.presetAdjustments?.adj3 ?? 25_000) / 100_000
  const innerWidth = Math.max(0, width - 2 * inset)
  const innerHeight = Math.max(0, height - 2 * inset)
  const innerPoint = (angle: number): [number, number] => {
    const [x, y] = presetEllipsePoint(innerWidth, innerHeight, angle)
    return [500 + (x - 500) * innerWidth / width, 500 + (y - 500) * innerHeight / height]
  }
  const arcCommands = (fromAngle: number, sweep: number, rx: number, ry: number, pointAt: (angle: number) => [number, number], direction: 0 | 1): string => {
    const [startX, startY] = pointAt(fromAngle)
    if (Math.abs(sweep) === fullTurn) {
      const [middleX, middleY] = pointAt(fromAngle + sweep / 2)
      return `A${rx} ${ry} 0 0 ${direction} ${middleX} ${middleY} A${rx} ${ry} 0 0 ${direction} ${startX} ${startY}`
    }
    const [endX, endY] = pointAt(fromAngle + sweep)
    return `A${rx} ${ry} 0 ${Math.abs(sweep) > fullTurn / 2 ? 1 : 0} ${direction} ${endX} ${endY}`
  }
  const [outerStartX, outerStartY] = presetEllipsePoint(width, height, startAngle)
  const [innerEndX, innerEndY] = innerPoint(endAngle)
  const outerArc = arcCommands(startAngle, sweepAngle, 500, 500, angle => presetEllipsePoint(width, height, angle), 1)
  const innerArc = arcCommands(endAngle, -sweepAngle, 500 * innerWidth / width, 500 * innerHeight / height, innerPoint, 0)
  return `M${outerStartX} ${outerStartY} ${outerArc} L${innerEndX} ${innerEndY} ${innerArc} Z`
}

function presetOutlinePath(element: PptxElement): string {
  if (element.geometry === 'leftBracket' || element.geometry === 'rightBracket' || element.geometry === 'bracketPair' || element.geometry === 'leftBrace' || element.geometry === 'rightBrace' || element.geometry === 'bracePair') return presetBracketBracePaths(element)[0]
  if (element.geometry === 'quadArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj2 = Math.max(0, Math.min(50000, element.presetAdjustments?.adj2 ?? 22500))
    const adj1 = Math.max(0, Math.min(2 * adj2, element.presetAdjustments?.adj1 ?? 22500))
    const adj3 = Math.max(0, Math.min((100000 - 2 * adj2) / 2, element.presetAdjustments?.adj3 ?? 22500))
    const x1 = shortSide * adj3 / 100000
    const dx2 = shortSide * adj2 / 100000
    const dx3 = shortSide * adj1 / 200000
    const x2 = width / 2 - dx2, x3 = width / 2 - dx3, x4 = width / 2 + dx3, x5 = width / 2 + dx2, x6 = width - x1
    const y2 = height / 2 - dx2, y3 = height / 2 - dx3, y4 = height / 2 + dx3, y5 = height / 2 + dx2, y6 = height - x1
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 500 L${x(x1)} ${y(y2)} L${x(x1)} ${y(y3)} L${x(x3)} ${y(y3)} L${x(x3)} ${y(x1)} L${x(x2)} ${y(x1)} L500 0 L${x(x5)} ${y(x1)} L${x(x4)} ${y(x1)} L${x(x4)} ${y(y3)} L${x(x6)} ${y(y3)} L${x(x6)} ${y(y2)} L1000 500 L${x(x6)} ${y(y5)} L${x(x6)} ${y(y4)} L${x(x4)} ${y(y4)} L${x(x4)} ${y(y6)} L${x(x5)} ${y(y6)} L500 1000 L${x(x2)} ${y(y6)} L${x(x3)} ${y(y6)} L${x(x3)} ${y(y4)} L${x(x1)} ${y(y4)} L${x(x1)} ${y(y5)} Z`
  }
  if (element.geometry === 'leftRightUpArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj2 = Math.max(0, Math.min(50000, element.presetAdjustments?.adj2 ?? 25000))
    const adj1 = Math.max(0, Math.min(2 * adj2, element.presetAdjustments?.adj1 ?? 25000))
    const adj3 = Math.max(0, Math.min((100000 - 2 * adj2) / 2, element.presetAdjustments?.adj3 ?? 25000))
    const x1 = shortSide * adj3 / 100000
    const dx2 = shortSide * adj2 / 100000
    const dx3 = shortSide * adj1 / 200000
    const x2 = width / 2 - dx2, x3 = width / 2 - dx3, x4 = width / 2 + dx3, x5 = width / 2 + dx2, x6 = width - x1
    const y2 = height - shortSide * adj2 / 50000, y4 = height - dx2, y3 = y4 - dx3, y5 = y4 + dx3
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 ${y(y4)} L${x(x1)} ${y(y2)} L${x(x1)} ${y(y3)} L${x(x3)} ${y(y3)} L${x(x3)} ${y(x1)} L${x(x2)} ${y(x1)} L500 0 L${x(x5)} ${y(x1)} L${x(x4)} ${y(x1)} L${x(x4)} ${y(y3)} L${x(x6)} ${y(y3)} L${x(x6)} ${y(y2)} L1000 ${y(y4)} L${x(x6)} 1000 L${x(x6)} ${y(y5)} L${x(x1)} ${y(y5)} L${x(x1)} 1000 Z`
  }
  if (element.geometry === 'bentUpArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(50000, element.presetAdjustments?.adj1 ?? 25000))
    const adj2 = Math.max(0, Math.min(50000, element.presetAdjustments?.adj2 ?? 25000))
    const adj3 = Math.max(0, Math.min(50000, element.presetAdjustments?.adj3 ?? 25000))
    const dx2 = shortSide * adj1 / 200000
    const x3 = width - shortSide * adj1 / 100000
    const x2 = x3 - dx2, x4 = x3 + dx2
    const y1 = shortSide * adj3 / 100000
    const y2 = height - shortSide * adj1 / 100000
    const x1 = width - shortSide * adj2 / 50000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 ${y(y2)} L${x(x2)} ${y(y2)} L${x(x2)} ${y(y1)} L${x(x1)} ${y(y1)} L${x(x3)} 0 L1000 ${y(y1)} L${x(x4)} ${y(y1)} V1000 H0 Z`
  }
  if (element.geometry === 'upArrow' || element.geometry === 'downArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 50000))
    const adj2 = Math.max(0, Math.min(100000 * height / shortSide, element.presetAdjustments?.adj2 ?? 50000))
    const halfHeadWidth = width * adj1 / 200000
    const x1 = width / 2 - halfHeadWidth, x2 = width / 2 + halfHeadWidth
    const arrowDepth = shortSide * adj2 / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    if (element.geometry === 'upArrow') {
      const y2 = arrowDepth
      return `M0 ${y(y2)} L500 0 L1000 ${y(y2)} L${x(x2)} ${y(y2)} V1000 H${x(x1)} V${y(y2)} Z`
    }
    const y1 = height - arrowDepth
    return `M0 ${y(y1)} L${x(x1)} ${y(y1)} V0 H${x(x2)} V${y(y1)} H1000 L500 1000 Z`
  }
  if (element.geometry === 'leftArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 50000))
    const adj2 = Math.max(0, Math.min(100000 * width / shortSide, element.presetAdjustments?.adj2 ?? 50000))
    const arrowLength = shortSide * adj2 / 100000
    const x2 = arrowLength
    const arrowHalfHeight = height * adj1 / 200000
    const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 500 L${x(x2)} 0 L${x(x2)} ${y(y1)} H1000 V${y(y2)} H${x(x2)} L${x(x2)} 1000 Z`
  }
  if (element.geometry === 'notchedRightArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 50000))
    const adj2 = Math.max(0, Math.min(100000 * width / shortSide, element.presetAdjustments?.adj2 ?? 50000))
    const arrowLength = shortSide * adj2 / 100000
    const x2 = width - arrowLength
    const arrowHalfHeight = height * adj1 / 200000
    const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight
    const x1 = arrowHalfHeight * arrowLength / (height / 2)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 ${y(y1)} L${x(x2)} ${y(y1)} L${x(x2)} 0 L1000 500 L${x(x2)} 1000 L${x(x2)} ${y(y2)} H0 L${x(x1)} 500 Z`
  }
  if (element.geometry === 'rightArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 50000))
    const adj2 = Math.max(0, Math.min(100000 * width / shortSide, element.presetAdjustments?.adj2 ?? 50000))
    const arrowLength = shortSide * adj2 / 100000
    const x1 = width - arrowLength
    const arrowHalfHeight = height * adj1 / 200000
    const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 ${y(y1)} L${x(x1)} ${y(y1)} L${x(x1)} 0 L1000 500 L${x(x1)} 1000 L${x(x1)} ${y(y2)} H0 Z`
  }
  if (element.geometry === 'upDownArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 50000))
    const adj2 = Math.max(0, Math.min(50000 * height / shortSide, element.presetAdjustments?.adj2 ?? 50000))
    const arrowHalfWidth = width * adj1 / 200000
    const x1 = width / 2 - arrowHalfWidth, x2 = width / 2 + arrowHalfWidth
    const arrowHeight = shortSide * adj2 / 100000
    const y2 = arrowHeight, y3 = height - arrowHeight
    const dy1 = x1 * y2 / (width / 2), y1 = y2 - dy1, y4 = y3 + dy1
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M500 0 L1000 ${y(y2)} L${x(x2)} ${y(y2)} L${x(x2)} ${y(y3)} L1000 ${y(y3)} L500 1000 L0 ${y(y3)} L${x(x1)} ${y(y3)} L${x(x1)} ${y(y2)} L0 ${y(y2)} Z`
  }
  if (element.geometry === 'leftRightArrow') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 50000))
    const adj2 = Math.max(0, Math.min(50000 * width / shortSide, element.presetAdjustments?.adj2 ?? 50000))
    const arrowWidth = shortSide * adj2 / 100000
    const x2 = arrowWidth, x3 = width - arrowWidth
    const arrowHalfHeight = height * adj1 / 200000
    const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight
    const dx1 = y1 * x2 / (height / 2), x1 = x2 - dx1, x4 = x3 + dx1
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 500 L${x(x2)} 0 L${x(x2)} ${y(y1)} L${x(x3)} ${y(y1)} L${x(x3)} 0 L1000 500 L${x(x3)} 1000 L${x(x3)} ${y(y2)} L${x(x2)} ${y(y2)} L${x(x2)} 1000 Z`
  }
  if (element.geometry === 'plus') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const inset = Math.min(width, height) * Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.25))
    const x1 = inset / width * 1000, x2 = 1000 - x1
    const y1 = inset / height * 1000, y2 = 1000 - y1
    return `M0 ${y1} L${x1} ${y1} L${x1} 0 L${x2} 0 L${x2} ${y1} H1000 V${y2} H${x2} V1000 H${x1} V${y2} H0 Z`
  }
  if (element.geometry === 'mathMinus') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const adj1 = Math.max(0, Math.min(100000, element.presetAdjustments?.adj1 ?? 23520))
    const halfBarWidth = width * 73490 / 200000
    const halfBarHeight = height * adj1 / 200000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(width / 2 - halfBarWidth)} ${y(height / 2 - halfBarHeight)} H${x(width / 2 + halfBarWidth)} V${y(height / 2 + halfBarHeight)} H${x(width / 2 - halfBarWidth)} Z`
  }
  if (element.geometry === 'mathEqual') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const adj1 = Math.max(0, Math.min(36745, element.presetAdjustments?.adj1 ?? 23520))
    const adj2 = Math.max(0, Math.min(100000 - 2 * adj1, element.presetAdjustments?.adj2 ?? 11760))
    const halfBarWidth = width * 73490 / 200000
    const halfGap = height * adj2 / 200000
    const barHeight = height * adj1 / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const left = x(width / 2 - halfBarWidth), right = x(width / 2 + halfBarWidth)
    const upperInner = y(height / 2 - halfGap), upperOuter = y(height / 2 - halfGap - barHeight)
    const lowerInner = y(height / 2 + halfGap), lowerOuter = y(height / 2 + halfGap + barHeight)
    return `M${left} ${upperOuter} H${right} V${upperInner} H${left} Z M${left} ${lowerInner} H${right} V${lowerOuter} H${left} Z`
  }
  if (element.geometry === 'mathPlus') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const adj1 = Math.max(0, Math.min(73490, element.presetAdjustments?.adj1 ?? 23520))
    const longHalfArm = width * 73490 / 200000
    const crossHalfArm = Math.min(width, height) * adj1 / 200000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const x1 = x(width / 2 - longHalfArm), x2 = x(width / 2 - crossHalfArm), x3 = x(width / 2 + crossHalfArm), x4 = x(width / 2 + longHalfArm)
    const y1 = y(height / 2 - height * 73490 / 200000), y2 = y(height / 2 - crossHalfArm), y3 = y(height / 2 + crossHalfArm), y4 = y(height / 2 + height * 73490 / 200000)
    return `M${x1} ${y2} L${x2} ${y2} L${x2} ${y1} L${x3} ${y1} L${x3} ${y2} L${x4} ${y2} L${x4} ${y3} L${x3} ${y3} L${x3} ${y4} L${x2} ${y4} L${x2} ${y3} L${x1} ${y3} Z`
  }
  if (element.geometry === 'mathDivide') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const adj1 = Math.max(1000, Math.min(36745, element.presetAdjustments?.adj1 ?? 23520))
    const maximumAdj3 = Math.min((73490 + adj1) / 4, 36745 * width / height)
    const adj3 = Math.max(1000, Math.min(maximumAdj3, element.presetAdjustments?.adj3 ?? 11760))
    const maximumAdj2 = 73490 - 4 * adj3 - adj1
    const adj2 = Math.max(0, Math.min(maximumAdj2, element.presetAdjustments?.adj2 ?? 5880))
    const dy1 = height * adj1 / 200000, gap = height * adj2 / 100000, radius = height * adj3 / 100000
    const y3 = height / 2 - dy1, y4 = height / 2 + dy1
    const y1 = y3 - gap, y2 = y1 + radius, y5 = height - y1
    const halfBarWidth = width * 73490 / 200000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const circle = (centerY: number) => `M${x(width / 2)} ${y(centerY - radius)} A${x(radius)} ${y(radius)} 0 1 1 ${x(width / 2)} ${y(centerY + radius)} A${x(radius)} ${y(radius)} 0 1 1 ${x(width / 2)} ${y(centerY - radius)} Z`
    return `${circle(y2)} ${circle(y5 - radius)} M${x(width / 2 - halfBarWidth)} ${y(y3)} H${x(width / 2 + halfBarWidth)} V${y(y4)} H${x(width / 2 - halfBarWidth)} Z`
  }
  if (element.geometry === 'mathMultiply') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const adj1 = Math.max(0, Math.min(51965, element.presetAdjustments?.adj1 ?? 23520))
    const angle = Math.atan2(height, width), tangent = Math.tan(angle)
    const widthAlongDiagonal = Math.hypot(width, height) * (1 - 0.51965)
    const thickness = Math.min(width, height) * adj1 / 100000
    const xM = Math.cos(angle) * widthAlongDiagonal / 2, yM = Math.sin(angle) * widthAlongDiagonal / 2
    const dx = Math.sin(angle) * thickness / 2, dy = Math.cos(angle) * thickness / 2
    const xA = xM - dx, yA = yM + dy, xB = xM + dx, yB = yM - dy
    const yC = (width / 2 - xB) * tangent + yB
    const xD = width - xB, xE = width - xA
    const xFE = tangent ? (height / 2 - yA) / tangent : 0
    const xF = xE - xFE, xL = xA + xFE
    const yG = height - yA, yH = height - yB, yI = height - yC
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(xA)} ${y(yA)} L${x(xB)} ${y(yB)} L500 ${y(yC)} L${x(xD)} ${y(yB)} L${x(xE)} ${y(yA)} L${x(xF)} 500 L${x(xE)} ${y(yG)} L${x(xD)} ${y(yH)} L500 ${y(yI)} L${x(xB)} ${y(yH)} L${x(xA)} ${y(yG)} L${x(xL)} 500 Z`
  }
  if (element.geometry === 'mathNotEqual') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const adj1 = Math.max(0, Math.min(50000, element.presetAdjustments?.adj1 ?? 23520))
    const angle = Math.max(4200000, Math.min(6600000, element.presetAdjustments?.adj2 ?? 6600000)) - 5400000
    const adj3 = Math.max(0, Math.min(100000 - 2 * adj1, element.presetAdjustments?.adj3 ?? 11760))
    const halfHeight = height / 2
    const dy1 = height * adj1 / 100000, dy2 = height * adj3 / 200000
    const x1 = width / 2 - width * 73490 / 200000, x8 = width / 2 + width * 73490 / 200000
    const y2 = halfHeight - dy2, y3 = halfHeight + dy2, y1 = y2 - dy1, y4 = y3 + dy1
    const xAdj = halfHeight * Math.tan(angle / 60000 * Math.PI / 180)
    const length = Math.hypot(xAdj, halfHeight)
    const barHalfWidth = length * dy1 / halfHeight / 2
    const x7 = width / 2 + xAdj - barHalfWidth
    const x6 = x7 - xAdj * y1 / halfHeight, x5 = x7 - xAdj * y2 / halfHeight
    const x4 = x7 - xAdj * y3 / halfHeight, x3 = x7 - xAdj * y4 / halfHeight
    const rx7 = x7 - barHalfWidth, rx6 = x6 - barHalfWidth
    const rx5 = x5 - barHalfWidth, rx4 = x4 - barHalfWidth, rx3 = x3 - barHalfWidth
    const dx7 = dy1 * halfHeight / length
    const dy3 = dy1 * xAdj / length
    const rx = angle > 0 ? x7 + dx7 : rx7
    const lx = angle > 0 ? x7 : rx7 - dx7
    const ry = angle > 0 ? dy3 : 0
    const ly = angle > 0 ? 0 : -dy3
    const dlx = width - rx, drx = width - lx, dly = height - ry, dry = height - ly
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(x1)} ${y(y1)} L${x(x6)} ${y(y1)} L${x(lx)} ${y(ly)} L${x(rx)} ${y(ry)} L${x(rx6)} ${y(y1)} L${x(x8)} ${y(y1)} L${x(x8)} ${y(y2)} L${x(rx5)} ${y(y2)} L${x(rx4)} ${y(y3)} L${x(x8)} ${y(y3)} L${x(x8)} ${y(y4)} L${x(rx3)} ${y(y4)} L${x(drx)} ${y(dry)} L${x(dlx)} ${y(dly)} L${x(x3)} ${y(y4)} L${x(x1)} ${y(y4)} L${x(x1)} ${y(y3)} L${x(x4)} ${y(y3)} L${x(x5)} ${y(y2)} L${x(x1)} ${y(y2)} Z`
  }
  if (element.geometry === 'snip1Rect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const inset = Math.min(width, height) * (element.presetAdjustment ?? 0.16667)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 0 H${x(width - inset)} L1000 ${y(inset)} V1000 H0 Z`
  }
  if (element.geometry === 'snip2DiagRect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const leftCut = shortSide * Math.max(0, Math.min(50000, element.presetAdjustments?.adj1 ?? 0)) / 100000
    const rightCut = shortSide * Math.max(0, Math.min(50000, element.presetAdjustments?.adj2 ?? 16667)) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(leftCut)} 0 L${x(width - rightCut)} 0 L1000 ${y(rightCut)} L1000 ${y(height - leftCut)} L${x(width - leftCut)} 1000 L${x(rightCut)} 1000 L0 ${y(height - rightCut)} L0 ${y(leftCut)} Z`
  }
  if (element.geometry === 'snip2SameRect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const topCut = shortSide * Math.max(0, Math.min(50000, element.presetAdjustments?.adj1 ?? 16667)) / 100000
    const bottomCut = shortSide * Math.max(0, Math.min(50000, element.presetAdjustments?.adj2 ?? 0)) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(topCut)} 0 L${x(width - topCut)} 0 L1000 ${y(topCut)} L1000 ${y(height - bottomCut)} L${x(width - bottomCut)} 1000 L${x(bottomCut)} 1000 L0 ${y(height - bottomCut)} L0 ${y(topCut)} Z`
  }
  if (element.geometry === 'round1Rect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const radius = Math.min(width, height) * (element.presetAdjustment ?? 0.16667)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 0 L${x(width - radius)} 0 A${x(radius)} ${y(radius)} 0 0 1 1000 ${y(radius)} L1000 1000 H0 Z`
  }
  if (element.geometry === 'round2DiagRect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const radius1 = shortSide * (element.presetAdjustments?.adj1 ?? 16667) / 100000
    const radius2 = shortSide * (element.presetAdjustments?.adj2 ?? 0) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(radius1)} 0 L${x(width - radius2)} 0 A${x(radius2)} ${y(radius2)} 0 0 1 1000 ${y(radius2)} L1000 ${y(height - radius1)} A${x(radius1)} ${y(radius1)} 0 0 1 ${x(width - radius1)} 1000 L${x(radius2)} 1000 A${x(radius2)} ${y(radius2)} 0 0 1 0 ${y(height - radius2)} L0 ${y(radius1)} A${x(radius1)} ${y(radius1)} 0 0 1 ${x(radius1)} 0 Z`
  }
  if (element.geometry === 'round2SameRect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const topRadius = shortSide * (element.presetAdjustments?.adj1 ?? 16667) / 100000
    const bottomRadius = shortSide * (element.presetAdjustments?.adj2 ?? 0) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(topRadius)} 0 L${x(width - topRadius)} 0 A${x(topRadius)} ${y(topRadius)} 0 0 1 1000 ${y(topRadius)} L1000 ${y(height - bottomRadius)} A${x(bottomRadius)} ${y(bottomRadius)} 0 0 1 ${x(width - bottomRadius)} 1000 L${x(bottomRadius)} 1000 A${x(bottomRadius)} ${y(bottomRadius)} 0 0 1 0 ${y(height - bottomRadius)} L0 ${y(topRadius)} A${x(topRadius)} ${y(topRadius)} 0 0 1 ${x(topRadius)} 0 Z`
  }
  if (element.geometry === 'snipRoundRect') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const radius = shortSide * (element.presetAdjustments?.adj1 ?? 16667) / 100000
    const cut = shortSide * (element.presetAdjustments?.adj2 ?? 16667) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(radius)} 0 L${x(width - cut)} 0 L1000 ${y(cut)} L1000 1000 H0 L0 ${y(radius)} A${x(radius)} ${y(radius)} 0 0 1 ${x(radius)} 0 Z`
  }
  if (element.geometry === 'cube') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const inset = Math.min(width, height) * (element.presetAdjustment ?? 0.25)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 ${y(inset)} L${x(inset)} 0 H1000 V${y(height - inset)} L${x(width - inset)} 1000 H0 Z`
  }
  if (element.geometry === 'frame') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const inset = shortSide * Math.max(0, Math.min(50000, element.presetAdjustments?.adj1 ?? 12500)) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 0 H1000 V1000 H0 Z M${x(inset)} ${y(inset)} V${y(height - inset)} H${x(width - inset)} V${y(inset)} Z`
  }
  if (element.geometry === 'halfFrame') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const maximumAdj2 = 100000 * width / shortSide
    const adj2 = Math.max(0, Math.min(maximumAdj2, element.presetAdjustments?.adj2 ?? 33333))
    const x1 = shortSide * adj2 / 100000
    const maximumAdj1 = 100000 * (height - height * x1 / width) / shortSide
    const adj1 = Math.max(0, Math.min(maximumAdj1, element.presetAdjustments?.adj1 ?? 33333))
    const y1 = shortSide * adj1 / 100000
    const dx2 = y1 * width / height
    const x2 = width - dx2
    const y2 = height - x1 * height / width
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 0 H1000 L${x(x2)} ${y(y1)} L${x(x1)} ${y(y1)} L${x(x1)} ${y(y2)} Z`
  }
  if (element.geometry === 'cornerTabs' || element.geometry === 'plaqueTabs' || element.geometry === 'squareTabs') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return ''
    const inset = Math.hypot(width, height) / 20
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const rightInset = x(width - inset), bottomInset = y(height - inset), topRight = y(inset), leftRight = x(inset)
    if (element.geometry === 'cornerTabs') {
      return `M0 0 L${leftRight} 0 L0 ${topRight} Z M0 ${bottomInset} L${leftRight} 1000 L0 1000 Z M${rightInset} 0 L1000 0 L1000 ${topRight} Z M1000 ${bottomInset} L1000 1000 L${rightInset} 1000 Z`
    }
    if (element.geometry === 'squareTabs') return `M0 0 L${leftRight} 0 L${leftRight} ${topRight} L0 ${topRight} Z M0 ${bottomInset} L${leftRight} ${bottomInset} L${leftRight} 1000 L0 1000 Z M${rightInset} 0 L1000 0 L1000 ${topRight} L${rightInset} ${topRight} Z M${rightInset} ${bottomInset} L1000 ${bottomInset} L1000 1000 L${rightInset} 1000 Z`
    return `M0 0 L${leftRight} 0 A${leftRight} ${topRight} 0 0 1 0 ${topRight} Z M0 ${bottomInset} A${leftRight} ${topRight} 0 0 1 ${leftRight} 1000 L0 1000 Z M1000 0 L1000 ${topRight} A${leftRight} ${topRight} 0 0 1 ${rightInset} 0 Z M${rightInset} 1000 A${leftRight} ${topRight} 0 0 1 1000 ${bottomInset} L1000 1000 Z`
  }
  if (element.geometry === 'plaque') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const radius = Math.min(width, height) * (element.presetAdjustment ?? 0.16667)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 ${y(radius)} A${x(radius)} ${y(radius)} 0 0 1 ${x(radius)} 0 L${x(width - radius)} 0 A${x(radius)} ${y(radius)} 0 0 1 1000 ${y(radius)} L1000 ${y(height - radius)} A${x(radius)} ${y(radius)} 0 0 1 ${x(width - radius)} 1000 L${x(radius)} 1000 A${x(radius)} ${y(radius)} 0 0 1 0 ${y(height - radius)} Z`
  }
  if (element.geometry === 'donut') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const inset = Math.min(width, height) * (element.presetAdjustment ?? 0.25)
    const innerRx = (width / 2 - inset) / width * 1000, innerRy = (height / 2 - inset) / height * 1000
    const innerPath = innerRx > 0 && innerRy > 0
      ? `M${inset / width * 1000} 500 A${innerRx} ${innerRy} 0 0 0 500 ${500 + innerRy} A${innerRx} ${innerRy} 0 0 0 ${1000 - inset / width * 1000} 500 A${innerRx} ${innerRy} 0 0 0 500 ${500 - innerRy} A${innerRx} ${innerRy} 0 0 0 ${inset / width * 1000} 500 Z`
      : ''
    return `M0 500 A500 500 0 0 1 500 0 A500 500 0 0 1 1000 500 A500 500 0 0 1 500 1000 A500 500 0 0 1 0 500 Z ${innerPath}`
  }
  if (['pie', 'chord', 'arc'].includes(element.geometry)) return presetAngleArcPaths(element)[0]
  if (element.geometry === 'blockArc') return presetBlockArcPath(element)
  if (element.geometry === 'pieWedge') return 'M0 1000 A1000 1000 0 0 1 1000 0 L1000 1000 Z'
  if (element.geometry === 'homePlate') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const cut = Math.min(width, height) * (element.presetAdjustment ?? 0.5)
    const x = (width - cut) / width * 1000
    return `M0 0 L${x} 0 L1000 500 L${x} 1000 L0 1000 Z`
  }
  if (element.geometry === 'nonIsoscelesTrapezoid') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const shortSide = Math.min(width, height)
    const leftTop = shortSide * (element.presetAdjustments?.adj1 ?? 25000) / 100000
    const rightTop = width - shortSide * (element.presetAdjustments?.adj2 ?? 25000) / 100000
    const x = (value: number) => value / width * 1000
    return `M0 1000 L${x(leftTop)} 0 L${x(rightTop)} 0 L1000 1000 Z`
  }
  if (element.geometry === 'heart') {
    // ECMA-376 Part 1 defines the heart with two symmetric cubic Bézier segments.
    const innerControlX = 500 + 1000 * 10 / 48, outerControlX = 500 + 1000 * 49 / 48
    const upperControlY = -1000 / 3
    return `M500 250 C${innerControlX} ${upperControlY} ${outerControlX} 250 500 1000 C${1000 - outerControlX} 250 ${1000 - innerControlX} ${upperControlY} 500 250 Z`
  }
  if (element.geometry === 'lightningBolt') {
    // ECMA-376 Part 1 defines the lightning bolt's eleven preset path points in a 21600-unit square.
    const scale = 1000 / 21600
    const points: [number, number][] = [[8472, 0], [12860, 6080], [11050, 6797], [16577, 12007], [14767, 12877], [21600, 21600], [10012, 14915], [12222, 13987], [5022, 9705], [7602, 8382], [0, 3890]]
    return points.map(([pointX, pointY], index) => (index ? 'L' : 'M') + pointX * scale + ' ' + pointY * scale).join(' ') + ' Z'
  }
  if (element.geometry === 'triangle') return `M0 1000 L${(element.presetAdjustment ?? 0.5) * 1000} 0 L1000 1000 Z`
  if (element.geometry === 'rtTriangle') return 'M0 1000 L0 0 L1000 1000 Z'
  if (element.geometry === 'diamond' || element.geometry === 'flowChartDecision') return 'M0 500 L500 0 L1000 500 L500 1000 Z'
  if (element.geometry === 'flowChartOffpageConnector') return 'M0 0 H1000 V800 L500 1000 L0 800 Z'
  if (element.geometry === 'flowChartDocument') {
    const scale = 1000 / 21600
    return `M0 0 H1000 V${17322 * scale} C${10800 * scale} ${17322 * scale} ${10800 * scale} ${23922 * scale} 0 ${20172 * scale} Z`
  }
  if (element.geometry === 'flowChartManualInput') return 'M0 200 L1000 0 V1000 H0 Z'
  if (element.geometry === 'flowChartPunchedCard') return 'M0 200 L200 0 H1000 V1000 H0 Z'
  if (element.geometry === 'flowChartOfflineStorage') return 'M0 0 H1000 L500 1000 Z M400 800 H600'
  if (element.geometry === 'flowChartMagneticDisk') return 'M0 166.6667 A500 166.6667 0 0 1 1000 166.6667 L1000 833.3333 A500 166.6667 0 0 1 0 833.3333 Z'
  if (element.geometry === 'flowChartMagneticDrum') {
    const inset = 1000 / 6
    return `M${inset} 0 H${1000 - inset} A${inset} 500 0 0 1 ${1000 - inset} 1000 L${inset} 1000 A${inset} 500 0 0 1 ${inset} 0 Z`
  }
  if (element.geometry === 'flowChartMagneticTape') {
    const angle = Math.atan2(element.height, element.width)
    const x = 500 + 500 * Math.cos(angle), y = 500 + 500 * Math.sin(angle)
    return `M500 1000 A500 500 0 0 1 0 500 A500 500 0 0 1 500 0 A500 500 0 0 1 1000 500 A500 500 0 0 1 ${x} ${y} L1000 ${y} L1000 1000 Z`
  }
  if (element.geometry === 'flowChartMultidocument') {
    const scale = 1000 / 21600
    const unit = (value: number) => value * scale
    return `M0 ${unit(20782)} C${unit(9298)} ${unit(23542)} ${unit(9298)} ${unit(18022)} ${unit(18595)} ${unit(18022)} L${unit(18595)} ${unit(3675)} H0 Z M${unit(1532)} ${unit(3675)} L${unit(1532)} ${unit(1815)} L${unit(20000)} ${unit(1815)} L${unit(20000)} ${unit(16252)} C${unit(19298)} ${unit(16252)} ${unit(18595)} ${unit(16352)} ${unit(18595)} ${unit(16352)} L${unit(18595)} ${unit(3675)} Z M${unit(2972)} ${unit(1815)} L${unit(2972)} 0 H1000 L1000 ${unit(14392)} C${unit(20800)} ${unit(14392)} ${unit(20000)} ${unit(14467)} ${unit(20000)} ${unit(14467)} L${unit(20000)} ${unit(1815)} Z`
  }
  if (element.geometry === 'flowChartOnlineStorage') {
    const inset = 1000 / 6
    return `M${inset} 0 H1000 A${inset} 500 0 0 0 1000 1000 L${inset} 1000 A${inset} 500 0 0 1 ${inset} 0 Z`
  }
  if (element.geometry === 'flowChartPunchedTape') return 'M0 100 A250 100 0 0 0 500 100 A250 100 0 0 1 1000 100 L1000 900 A250 100 0 0 0 500 900 A250 100 0 0 1 0 900 Z'
  if (element.geometry === 'flowChartManualOperation') return 'M0 0 H1000 L800 1000 L200 1000 Z'
  if (element.geometry === 'flowChartInputOutput') return 'M0 1000 L200 0 H1000 L800 1000 Z'
  if (element.geometry === 'flowChartPreparation') return 'M0 500 L200 0 H800 L1000 500 L800 1000 H200 Z'
  if (element.geometry === 'flowChartTerminator') {
    const radiusX = 3475 / 21600 * 1000
    return `M${radiusX} 0 H${1000 - radiusX} A${radiusX} 500 0 0 1 1000 500 L${radiusX} 1000 A${radiusX} 500 0 0 1 0 500 Z`
  }
  if (element.geometry === 'flowChartPredefinedProcess') return 'M0 0 H1000 V1000 H0 Z M125 0 V1000 M875 0 V1000'
  if (element.geometry === 'flowChartInternalStorage') return 'M0 0 H1000 V1000 H0 Z M125 0 V1000 M0 125 H1000'
  if (element.geometry === 'flowChartSort') return 'M0 500 L500 0 L1000 500 L500 1000 Z M0 500 H1000'
  if (element.geometry === 'flowChartCollate') return 'M0 0 H1000 L500 500 L1000 1000 H0 L500 500 Z'
  if (element.geometry === 'flowChartExtract') return 'M0 1000 L500 0 L1000 1000 Z'
  if (element.geometry === 'flowChartMerge') return 'M0 0 L1000 0 L500 1000 Z'
  if (element.geometry === 'trapezoid') {
    const cut = element.width > 0 ? Math.min(element.width, element.height) * (element.presetAdjustment ?? 0.25) / element.width * 1000 : 0
    return `M0 1000 L${cut} 0 L${1000 - cut} 0 L1000 1000 Z`
  }
  if (element.geometry === 'octagon') {
    const shortSide = Math.min(element.width, element.height)
    const adjustment = element.presetAdjustment ?? 0.29289
    const cutX = element.width > 0 ? shortSide * adjustment / element.width * 1000 : 0
    const cutY = element.height > 0 ? shortSide * adjustment / element.height * 1000 : 0
    return `M0 ${cutY} L${cutX} 0 H${1000 - cutX} L1000 ${cutY} V${1000 - cutY} L${1000 - cutX} 1000 H${cutX} L0 ${1000 - cutY} Z`
  }
  if (element.geometry === 'pentagon') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const halfWidth = width * (element.presetAdjustments?.hf ?? 105146) / 200000
    const halfHeight = height * (element.presetAdjustments?.vf ?? 110557) / 200000
    const dx1 = halfWidth * Math.cos(Math.PI / 10), dx2 = halfWidth * Math.cos(17 * Math.PI / 10)
    const dy1 = halfHeight * Math.sin(Math.PI / 10), dy2 = halfHeight * Math.sin(17 * Math.PI / 10)
    const centerY = height * (element.presetAdjustments?.vf ?? 110557) / 200000
    const pointX = (value: number) => value / width * 1000, pointY = (value: number) => value / height * 1000
    return `M${pointX(width / 2 - dx1)} ${pointY(centerY - dy1)} L500 0 L${pointX(width / 2 + dx1)} ${pointY(centerY - dy1)} L${pointX(width / 2 + dx2)} ${pointY(centerY - dy2)} L${pointX(width / 2 - dx2)} ${pointY(centerY - dy2)} Z`
  }
  if (element.geometry === 'heptagon') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const halfWidth = width / 2 * (element.presetAdjustments?.hf ?? 102572) / 100000
    const halfHeight = height / 2 * (element.presetAdjustments?.vf ?? 105210) / 100000
    const dx1 = halfWidth * 97493 / 100000, dx2 = halfWidth * 78183 / 100000, dx3 = halfWidth * 43388 / 100000
    const y1 = halfHeight - halfHeight * 62349 / 100000
    const y2 = halfHeight + halfHeight * 22252 / 100000
    const y3 = halfHeight + halfHeight * 90097 / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(width / 2 - dx1)} ${y(y2)} L${x(width / 2 - dx2)} ${y(y1)} L500 0 L${x(width / 2 + dx2)} ${y(y1)} L${x(width / 2 + dx1)} ${y(y2)} L${x(width / 2 + dx3)} ${y(y3)} L${x(width / 2 - dx3)} ${y(y3)} Z`
  }
  if (element.geometry === 'decagon') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const halfHeight = height / 2 * (element.presetAdjustments?.vf ?? 105146) / 100000
    const dx1 = width / 2 * Math.cos(Math.PI / 5), dx2 = width / 2 * Math.cos(2 * Math.PI / 5)
    const dy1 = halfHeight * Math.sin(2 * Math.PI / 5), dy2 = halfHeight * Math.sin(Math.PI / 5)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 500 L${x(width / 2 - dx1)} ${y(height / 2 - dy2)} L${x(width / 2 - dx2)} ${y(height / 2 - dy1)} L${x(width / 2 + dx2)} ${y(height / 2 - dy1)} L${x(width / 2 + dx1)} ${y(height / 2 - dy2)} L1000 500 L${x(width / 2 + dx1)} ${y(height / 2 + dy2)} L${x(width / 2 + dx2)} ${y(height / 2 + dy1)} L${x(width / 2 - dx2)} ${y(height / 2 + dy1)} L${x(width / 2 - dx1)} ${y(height / 2 + dy2)} Z`
  }
  if (element.geometry === 'star5') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const horizontalFactor = element.presetAdjustments?.hf ?? 105146, verticalFactor = element.presetAdjustments?.vf ?? 110557
    const halfWidth = width / 2 * horizontalFactor / 100000, halfHeight = height / 2 * verticalFactor / 100000
    const centerY = height / 2 * verticalFactor / 100000
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.19098)) * 2
    const innerHalfWidth = halfWidth * innerScale, innerHalfHeight = halfHeight * innerScale
    const dx1 = halfWidth * Math.cos(Math.PI / 10), dx2 = halfWidth * Math.cos(17 * Math.PI / 10)
    const dy1 = halfHeight * Math.sin(Math.PI / 10), dy2 = halfHeight * Math.sin(17 * Math.PI / 10)
    const sdx1 = innerHalfWidth * Math.cos(19 * Math.PI / 10), sdx2 = innerHalfWidth * Math.cos(3 * Math.PI / 10)
    const sdy1 = innerHalfHeight * Math.sin(3 * Math.PI / 10), sdy2 = innerHalfHeight * Math.sin(19 * Math.PI / 10)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(width / 2 - dx1)} ${y(centerY - dy1)} L${x(width / 2 - sdx2)} ${y(centerY - sdy1)} L500 0 L${x(width / 2 + sdx2)} ${y(centerY - sdy1)} L${x(width / 2 + dx1)} ${y(centerY - dy1)} L${x(width / 2 + sdx1)} ${y(centerY - sdy2)} L${x(width / 2 + dx2)} ${y(centerY - dy2)} L500 ${y(centerY + innerHalfHeight)} L${x(width / 2 - dx2)} ${y(centerY - dy2)} L${x(width / 2 - sdx1)} ${y(centerY - sdy2)} Z`
  }
  if (element.geometry === 'star4') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.125)) * 2
    const innerHalfWidth = width / 2 * innerScale, innerHalfHeight = height / 2 * innerScale
    const insetX = innerHalfWidth * Math.cos(Math.PI / 4), insetY = innerHalfHeight * Math.sin(Math.PI / 4)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 500 L${x(width / 2 - insetX)} ${y(height / 2 - insetY)} L500 0 L${x(width / 2 + insetX)} ${y(height / 2 - insetY)} L1000 500 L${x(width / 2 + insetX)} ${y(height / 2 + insetY)} L500 1000 L${x(width / 2 - insetX)} ${y(height / 2 + insetY)} Z`
  }
  if (element.geometry === 'star6') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const halfWidth = width / 2 * (element.presetAdjustments?.hf ?? 115470) / 100000
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.28868)) * 2
    const innerHalfWidth = halfWidth * innerScale, innerHalfHeight = height / 2 * innerScale
    const dx = halfWidth * Math.cos(Math.PI / 6), dy = innerHalfHeight * Math.sin(Math.PI / 3)
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M${x(width / 2 - dx)} ${y(height / 4)} L${x(width / 2 - innerHalfWidth / 2)} ${y(height / 2 - dy)} L500 0 L${x(width / 2 + innerHalfWidth / 2)} ${y(height / 2 - dy)} L${x(width / 2 + dx)} ${y(height / 4)} L${x(width / 2 + innerHalfWidth)} 500 L${x(width / 2 + dx)} ${y(height * 3 / 4)} L${x(width / 2 + innerHalfWidth / 2)} ${y(height / 2 + dy)} L500 1000 L${x(width / 2 - innerHalfWidth / 2)} ${y(height / 2 + dy)} L${x(width / 2 - dx)} ${y(height * 3 / 4)} L${x(width / 2 - innerHalfWidth)} 500 Z`
  }
  if (element.geometry === 'star7') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const horizontalFactor = element.presetAdjustments?.hf ?? 102572, verticalFactor = element.presetAdjustments?.vf ?? 105210
    const outerHalfWidth = width / 2 * horizontalFactor / 100000, outerHalfHeight = height / 2 * verticalFactor / 100000
    const centerX = width / 2, centerY = height / 2 * verticalFactor / 100000
    const outerDx1 = outerHalfWidth * 0.97493, outerDx2 = outerHalfWidth * 0.78183, outerDx3 = outerHalfWidth * 0.43388
    const outerDy1 = outerHalfHeight * 0.62349, outerDy2 = outerHalfHeight * 0.22252, outerDy3 = outerHalfHeight * 0.90097
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.34601)) * 2
    const innerHalfWidth = outerHalfWidth * innerScale, innerHalfHeight = outerHalfHeight * innerScale
    const innerDx1 = innerHalfWidth * 0.97493, innerDx2 = innerHalfWidth * 0.78183, innerDx3 = innerHalfWidth * 0.43388
    const innerDy1 = innerHalfHeight * 0.90097, innerDy2 = innerHalfHeight * 0.22252, innerDy3 = innerHalfHeight * 0.62349
    const x1 = centerX - outerDx1, x2 = centerX - outerDx2, x3 = centerX - outerDx3
    const x4 = centerX + outerDx3, x5 = centerX + outerDx2, x6 = centerX + outerDx1
    const y1 = centerY - outerDy1, y2 = centerY + outerDy2, y3 = centerY + outerDy3
    const sx1 = centerX - innerDx1, sx2 = centerX - innerDx2, sx3 = centerX - innerDx3
    const sx4 = centerX + innerDx3, sx5 = centerX + innerDx2, sx6 = centerX + innerDx1
    const sy1 = centerY - innerDy1, sy2 = centerY - innerDy2, sy3 = centerY + innerDy3, sy4 = centerY + innerHalfHeight
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const points: [number, number][] = [[x1, y2], [sx1, sy2], [x2, y1], [sx3, sy1], [centerX, 0], [sx4, sy1], [x5, y1], [sx6, sy2], [x6, y2], [sx5, sy3], [x4, y3], [centerX, sy4], [x3, y3], [sx2, sy3]]
    return points.map(([pointX, pointY], index) => (index ? 'L' : 'M') + x(pointX) + ' ' + y(pointY)).join(' ') + ' Z'
  }
  if (element.geometry === 'star8') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.375)) * 2
    const innerHalfWidth = width / 2 * innerScale, innerHalfHeight = height / 2 * innerScale
    const outerDx = width / 2 * Math.cos(Math.PI / 4), outerDy = height / 2 * Math.sin(Math.PI / 4)
    const innerDx1 = innerHalfWidth * 0.92388, innerDx2 = innerHalfWidth * 0.38268
    const innerDy1 = innerHalfHeight * 0.92388, innerDy2 = innerHalfHeight * 0.38268
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    return `M0 500 L${x(width / 2 - innerDx1)} ${y(height / 2 - innerDy2)} L${x(width / 2 - outerDx)} ${y(height / 2 - outerDy)} L${x(width / 2 - innerDx2)} ${y(height / 2 - innerDy1)} L500 0 L${x(width / 2 + innerDx2)} ${y(height / 2 - innerDy1)} L${x(width / 2 + outerDx)} ${y(height / 2 - outerDy)} L${x(width / 2 + innerDx1)} ${y(height / 2 - innerDy2)} L1000 500 L${x(width / 2 + innerDx1)} ${y(height / 2 + innerDy2)} L${x(width / 2 + outerDx)} ${y(height / 2 + outerDy)} L${x(width / 2 + innerDx2)} ${y(height / 2 + innerDy1)} L500 1000 L${x(width / 2 - innerDx2)} ${y(height / 2 + innerDy1)} L${x(width / 2 - outerDx)} ${y(height / 2 + outerDy)} L${x(width / 2 - innerDx1)} ${y(height / 2 + innerDy2)} Z`
  }
  if (element.geometry === 'star10') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const outerHalfWidth = width / 2 * (element.presetAdjustments?.hf ?? 105146) / 100000
    const outerDx1 = outerHalfWidth * 0.95106, outerDx2 = outerHalfWidth * 0.58779
    const outerDy1 = height / 2 * 0.80902, outerDy2 = height / 2 * 0.30902
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.42533)) * 2
    const innerHalfWidth = outerHalfWidth * innerScale, innerHalfHeight = height / 2 * innerScale
    const innerDx1 = innerHalfWidth * 0.80902, innerDx2 = innerHalfWidth * 0.30902
    const innerDy1 = innerHalfHeight * 0.95106, innerDy2 = innerHalfHeight * 0.58779
    const centerX = width / 2, centerY = height / 2
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const points: [number, number][] = [
      [centerX - outerDx1, centerY - outerDy2], [centerX - innerDx1, centerY - innerDy2],
      [centerX - outerDx2, centerY - outerDy1], [centerX - innerDx2, centerY - innerDy1], [centerX, 0],
      [centerX + innerDx2, centerY - innerDy1], [centerX + outerDx2, centerY - outerDy1], [centerX + innerDx1, centerY - innerDy2],
      [centerX + outerDx1, centerY - outerDy2], [centerX + innerHalfWidth, centerY], [centerX + outerDx1, centerY + outerDy2],
      [centerX + innerDx1, centerY + innerDy2], [centerX + outerDx2, centerY + outerDy1], [centerX + innerDx2, centerY + innerDy1],
      [centerX, height], [centerX - innerDx2, centerY + innerDy1], [centerX - outerDx2, centerY + outerDy1],
      [centerX - innerDx1, centerY + innerDy2], [centerX - outerDx1, centerY + outerDy2], [centerX - innerHalfWidth, centerY],
    ]
    return points.map(([pointX, pointY], index) => (index ? 'L' : 'M') + x(pointX) + ' ' + y(pointY)).join(' ') + ' Z'
  }
  if (element.geometry === 'star12') {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const centerX = width / 2, centerY = height / 2
    const outerDx = centerX * Math.cos(Math.PI / 6), outerDy = centerY * Math.sin(Math.PI / 3)
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.375)) * 2
    const innerHalfWidth = centerX * innerScale, innerHalfHeight = centerY * innerScale
    const innerDx1 = innerHalfWidth * Math.cos(Math.PI / 12), innerDx2 = innerHalfWidth * Math.cos(Math.PI / 4), innerDx3 = innerHalfWidth * Math.cos(5 * Math.PI / 12)
    const innerDy1 = innerHalfHeight * Math.sin(5 * Math.PI / 12), innerDy2 = innerHalfHeight * Math.sin(Math.PI / 4), innerDy3 = innerHalfHeight * Math.sin(Math.PI / 12)
    const x1 = centerX - outerDx, x3 = width * 3 / 4, x4 = centerX + outerDx
    const y1 = centerY - outerDy, y3 = height * 3 / 4, y4 = centerY + outerDy
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const points: [number, number][] = [
      [0, centerY], [centerX - innerDx1, centerY - innerDy3], [x1, height / 4], [centerX - innerDx2, centerY - innerDy2],
      [width / 4, y1], [centerX - innerDx3, centerY - innerDy1], [centerX, 0], [centerX + innerDx3, centerY - innerDy1],
      [x3, y1], [centerX + innerDx2, centerY - innerDy2], [x4, height / 4], [centerX + innerDx1, centerY - innerDy3],
      [width, centerY], [centerX + innerDx1, centerY + innerDy3], [x4, y3], [centerX + innerDx2, centerY + innerDy2],
      [x3, y4], [centerX + innerDx3, centerY + innerDy1], [centerX, height], [centerX - innerDx3, centerY + innerDy1],
      [width / 4, y4], [centerX - innerDx2, centerY + innerDy2], [x1, y3], [centerX - innerDx1, centerY + innerDy3],
    ]
    return points.map(([pointX, pointY], index) => (index ? 'L' : 'M') + x(pointX) + ' ' + y(pointY)).join(' ') + ' Z'
  }
  if (['star16', 'star24', 'star32'].includes(element.geometry)) {
    const width = element.width, height = element.height
    if (width <= 0 || height <= 0) return 'M0 0 H1000 V1000 H0 Z'
    const pointCount = Number(element.geometry.slice(4))
    const innerScale = Math.max(0, Math.min(0.5, element.presetAdjustment ?? 0.375)) * 2
    const centerX = width / 2, centerY = height / 2
    const guideFactor = (value: number) => Math.round(value * 100000) / 100000
    const x = (value: number) => value / width * 1000, y = (value: number) => value / height * 1000
    const points: [number, number][] = []
    for (let index = 0; index < pointCount * 2; index++) {
      const angle = Math.PI + index * Math.PI / pointCount
      const radius = index % 2 === 0 ? 1 : innerScale
      const usesExactTrig = element.geometry === 'star24' && index % 2 === 0 || element.geometry === 'star32' && index % 8 === 0
      const factor = (value: number) => usesExactTrig ? value : guideFactor(value)
      points.push([centerX + centerX * radius * factor(Math.cos(angle)), centerY + centerY * radius * factor(Math.sin(angle))])
    }
    return points.map(([pointX, pointY], index) => (index ? 'L' : 'M') + x(pointX) + ' ' + y(pointY)).join(' ') + ' Z'
  }
  if (element.geometry === 'dodecagon') {
    const scale = 1000 / 21600
    const x1 = 2894 * scale, x2 = 7906 * scale, x3 = 13694 * scale, x4 = 18706 * scale
    return `M0 ${x2} L${x1} ${x1} L${x2} 0 H${x3} L${x4} ${x1} L1000 ${x2} V${x3} L${x4} ${x4} L${x3} 1000 H${x2} L${x1} ${x4} L0 ${x3} Z`
  }
  if (element.geometry === 'parallelogram' || element.geometry === 'chevron' || element.geometry === 'hexagon') {
    const adjustment = element.presetAdjustment ?? (element.geometry === 'parallelogram' ? 0.25 : element.geometry === 'hexagon' ? 0.25 : 0.5)
    const shortSide = Math.min(element.width, element.height)
    const cut = shortSide > 0 && element.width > 0 ? shortSide * adjustment / element.width * 1000 : 0
    if (element.geometry === 'parallelogram') return `M0 1000 L${cut} 0 H1000 L${1000 - cut} 1000 Z`
    if (element.geometry === 'chevron') return `M0 0 H${1000 - cut} L1000 500 L${1000 - cut} 1000 H0 L${cut} 500 Z`
    const dy = 0.211325 * 1000
    return `M0 500 L${cut} ${dy} H${1000 - cut} L1000 500 L${1000 - cut} ${1000 - dy} H${cut} Z`
  }
  if (['ellipse', 'flowChartConnector', 'flowChartSummingJunction', 'flowChartOr'].includes(element.geometry)) {
    const ellipse = 'M0 500 A500 500 0 1 0 1000 500 A500 500 0 1 0 0 500 Z'
    if (element.geometry === 'flowChartSummingJunction') {
      const inset = (1 - Math.SQRT1_2) * 500
      return `${ellipse} M${inset} ${inset} L${1000 - inset} ${1000 - inset} M${1000 - inset} ${inset} L${inset} ${1000 - inset}`
    }
    if (element.geometry === 'flowChartOr') return `${ellipse} M500 0 V1000 M0 500 H1000`
    return ellipse
  }
  if (element.geometry === 'flowChartDelay') return 'M0 0 H500 A500 500 0 0 1 500 1000 H0 Z'
  if (element.geometry === 'flowChartDisplay') {
    const inset = 1000 / 6
    return `M0 500 L${inset} 0 H${1000 - inset} A${inset} 500 0 0 1 ${1000 - inset} 1000 L${inset} 1000 Z`
  }
  if (element.geometry !== 'roundRect' && element.geometry !== 'flowChartAlternateProcess') return 'M0 0 H1000 V1000 H0 Z'
  const [x, y] = roundRectRadii(element)
  if (!x || !y) return 'M0 0 H1000 V1000 H0 Z'
  return `M0 ${y} A${x} ${y} 0 0 1 ${x} 0 H${1000 - x} A${x} ${y} 0 0 1 1000 ${y} V${1000 - y} A${x} ${y} 0 0 1 ${1000 - x} 1000 H${x} A${x} ${y} 0 0 1 0 ${1000 - y} Z`
}

function customGeometryFill(element: PptxElement, path: NonNullable<PptxElement['customPaths']>[number]): string {
  if (!path.fill || !animationOnOff(element, 'fill.on')) return 'none'
  const animatedFill = animationColor(element, 'fillcolor')
  const brightness = customPathBrightness(path.fillMode)
  if (!animatedFill && isPatternFillCss(element.fill)) return `url(#${customGeometryPatternId(element, path.fillMode)})`
  if (!animatedFill && /^(?:linear|radial)-gradient\(/i.test(element.fill)) return `url(#${customGradientId(element, 'fill', path.fillMode)})`
  return applyColorOpacity(animatedFill || element.fill, animationOpacity(element, 'fill.opacity'), brightness)
}

function customGeometryFillOpacity(element: PptxElement, path: NonNullable<PptxElement['customPaths']>[number]): number | undefined {
  if (!path.fill || !animationOnOff(element, 'fill.on') || animationColor(element, 'fillcolor')) return undefined
  return isPatternFillCss(element.fill) ? animationOpacity(element, 'fill.opacity') : undefined
}

function customPathBrightness(mode: NonNullable<PptxElement['customPaths']>[number]['fillMode']): number {
  // OOXML names the modes; use LibreOffice's 20%/40% blends pending Office screenshot calibration.
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathfillmodevalues?view=openxml-3.0.1
  // https://github.com/LibreOffice/core/blob/master/svx/source/customshapes/EnhancedCustomShape2d.cxx
  return mode ? { lighten: 0.4, lightenLess: 0.2, darken: -0.4, darkenLess: -0.2 }[mode] : 0
}

function isPatternFillCss(value: string): boolean {
  return /repeating-linear-gradient\(/i.test(value) || /\) 0 0 \/ \d+(?:\.\d+)?px \d+(?:\.\d+)?px,/i.test(value)
}

function customGeometryStroke(element: PptxElement, path: NonNullable<PptxElement['customPaths']>[number]): string {
  if (!path.stroke || !animationOnOff(element, 'stroke.on')) return 'none'
  const animatedStroke = animationColor(element, 'stroke.color')
  const strokeWidth = animationStrokeWidth(element)
  if (strokeWidth > 0 && !animatedStroke && element.strokePattern) return `url(#${customGeometryStrokePatternId(element)})`
  if (path.stroke && strokeWidth > 0 && !animatedStroke && /^(?:linear|radial)-gradient\(/i.test(element.stroke)) return `url(#${customGradientId(element, 'stroke')})`
  const stroke = applyColorOpacity(animatedStroke || element.stroke, animationOpacity(element, 'stroke.opacity'))
  return path.stroke && strokeWidth > 0 && stroke !== 'transparent' && !stroke.includes('gradient') ? stroke : 'none'
}

function customGradientId(element: PptxElement, kind: 'fill' | 'stroke', mode?: NonNullable<PptxElement['customPaths']>[number]['fillMode']): string {
  return `${customGradientScope}-${kind}-${element.id.replace(/[^a-zA-Z0-9_-]/g, '_')}${mode ? `-${mode}` : ''}`
}

function customGeometryPatternModes(element: PptxElement) {
  if (!isPatternFillCss(element.fill)) return []
  return [...new Set([undefined, ...(element.customPaths?.flatMap(path => path.fillMode ? [path.fillMode] : []) || [])])]
}

function customGeometryPatternId(element: PptxElement, mode?: NonNullable<PptxElement['customPaths']>[number]['fillMode']): string {
  return `${customGradientScope}-pattern-${element.id.replace(/[^a-zA-Z0-9_-]/g, '_')}${mode ? `-${mode}` : ''}`
}

function customGeometryStrokePatternId(element: PptxElement): string {
  return `${customGradientScope}-stroke-pattern-${element.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
}

function customGradientFocusClipId(element: PptxElement, pathIndex: number): string {
  return `${customGradientScope}-fill-focus-${element.id.replace(/[^a-zA-Z0-9_-]/g, '_')}-${pathIndex}`
}

function splitGradientStops(value: string): string[] {
  const stops: string[] = []
  let depth = 0, start = 0
  for (let index = 0; index < value.length; index++) {
    if (value[index] === '(') depth++
    else if (value[index] === ')') depth--
    else if (value[index] === ',' && depth === 0) { stops.push(value.slice(start, index).trim()); start = index + 1 }
  }
  stops.push(value.slice(start).trim())
  return stops
}

const radialGradientBackgroundCache = new Map<string, string>()

function radialGradientBackground(gradient: CustomSvgRadialGradient): string {
  const key = JSON.stringify(gradient)
  const cached = radialGradientBackgroundCache.get(key)
  if (cached) return cached
  const stops = gradient.stops.map(stop => `<stop offset="${stop.offset}" stop-color="${stop.color}" stop-opacity="${stop.opacity}"/>`).join('')
  const transform = gradient.gradientTransform ? ` gradientTransform="${gradient.gradientTransform}"` : ''
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" preserveAspectRatio="none"><defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="${gradient.cx}" cy="${gradient.cy}" r="${gradient.r}" fx="${gradient.fx}" fy="${gradient.fy}" fr="${gradient.fr}"${transform}${gradient.colorInterpolation ? ` color-interpolation="${gradient.colorInterpolation}"` : ''}>${stops}</radialGradient></defs><rect width="1000" height="1000" fill="url(#g)"/></svg>`
  const background = `url("data:image/svg+xml,${encodeURIComponent(svg)}") center / 100% 100% no-repeat`
  if (radialGradientBackgroundCache.size >= 128) radialGradientBackgroundCache.delete(radialGradientBackgroundCache.keys().next().value!)
  radialGradientBackgroundCache.set(key, background)
  return background
}

function customSvgGradient(element: PptxElement, kind: 'fill' | 'stroke', mode?: NonNullable<PptxElement['customPaths']>[number]['fillMode']): CustomSvgGradient | undefined {
  const css = applyColorOpacity(kind === 'fill' ? element.fill : element.stroke, 1, customPathBrightness(mode))
  const opacity = kind === 'fill' ? animationOpacity(element, 'fill.opacity') : 1
  const layers = splitGradientStops(css)
  const linear = /^linear-gradient\(\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+))deg\s*,([\s\S]*)\)$/i.exec(layers.find(layer => layer.startsWith('linear-gradient(')) || '')
  const radialLayer = layers.find(layer => /^radial-gradient\(\s*(?:circle|ellipse)(?:\s+at\s+[-+\d.]+%\s+[-+\d.]+%)?\s*,/i.test(layer)) || ''
  const radial = /^radial-gradient\(\s*(circle|ellipse)(?:\s+at\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))%\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))%)?\s*,([\s\S]*)\)$/i.exec(radialLayer)
  if (!linear && !radial) return undefined
  const stops = splitGradientStops((linear ? linear[2] : radial![4])!).map((stop): CustomSvgGradientStop | undefined => {
    const match = /^(.*)\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))%$/.exec(stop)
    if (!match) return undefined
    const color = match[1]!.trim()
    const rgba = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*(?:,\s*([\d.]+)\s*)?\)$/i.exec(color)
    const hex = /^#([\da-f]{6})$/i.exec(color)
    if (color === 'transparent') return { offset: `${Math.max(0, Math.min(Number(match[2]), 100))}%`, color: '#000000', opacity: 0 }
    if (!rgba && !hex) return undefined
    return {
      offset: `${Math.max(0, Math.min(Number(match[2]), 100))}%`,
      color: rgba ? `rgb(${rgba[1]}, ${rgba[2]}, ${rgba[3]})` : color,
      opacity: (rgba && rgba[4] !== undefined ? Math.max(0, Math.min(Number(rgba[4]), 1)) : 1) * opacity,
    }
  })
  if (stops.length < 2 || stops.some(stop => !stop)) return undefined

  if (radial) {
    const focusEllipseLayer = layers.find(layer => /^radial-gradient\(\s*ellipse\s+closest-side\s*,[\s\S]*\)\s+left\s+[-+\d.]+%\s+top\s+[-+\d.]+%\s*\/\s*[-+\d.]+%\s+[-+\d.]+%\s+no-repeat$/i.test(layer))
    const focusEllipse = focusEllipseLayer && /^radial-gradient\(\s*ellipse\s+closest-side\s*,([\s\S]*)\)\s+left\s+([-+\d.]+)%\s+top\s+([-+\d.]+)%\s*\/\s*([-+\d.]+)%\s+([-+\d.]+)%\s+no-repeat$/i.exec(focusEllipseLayer)
    const focusRectLayer = layers.find(layer => /^linear-gradient\([\s\S]*\)\s+left\s+[-+\d.]+%\s+top\s+[-+\d.]+%\s*\/\s*[-+\d.]+%\s+[-+\d.]+%\s+no-repeat$/i.test(layer))
    const focusRect = focusRectLayer && /^linear-gradient\(([\s\S]*)\)\s+left\s+([-+\d.]+)%\s+top\s+([-+\d.]+)%\s*\/\s*([-+\d.]+)%\s+([-+\d.]+)%\s+no-repeat$/i.exec(focusRectLayer)
    const focus = focusEllipse || focusRect
    const focusColor = focusEllipse
      ? splitGradientStops(focusEllipse[1]!)[0]?.replace(/\s+[-+\d.]+%$/, '').trim()
      : focusRect ? splitGradientStops(focusRect[1]!)[0]?.trim() : undefined
    const cx = radial[2] === undefined ? 500 : Number(radial[2]) * 10
    const cy = radial[3] === undefined ? 500 : Number(radial[3]) * 10
    const left = focus ? Number(focus[2]) : 0
    const top = focus ? Number(focus[3]) : 0
    const width = focus ? Number(focus[4]) : 0
    const height = focus ? Number(focus[5]) : 0
    let gradientTransform: string | undefined
    let fx: number | undefined, fy: number | undefined, fr: number | undefined
    let focusOverlay: CustomSvgGradientFocus | undefined
    if (element.width > 0 && element.height > 0) {
      const circleRadius = radial[1]!.toLowerCase() === 'circle'
        ? Math.max(...[0, 1000].flatMap(x => [0, 1000].map(y => Math.hypot((x - cx) * element.width / 1000, (y - cy) * element.height / 1000))))
        : undefined
      const radiusX = circleRadius === undefined ? Math.max(cx, 1000 - cx) : circleRadius * 1000 / element.width
      const radiusY = circleRadius === undefined ? Math.max(cy, 1000 - cy) : circleRadius * 1000 / element.height
      const scaleX = radiusX / 500, scaleY = radiusY / 500
      gradientTransform = `translate(${cx} ${cy}) scale(${scaleX} ${scaleY}) translate(${-cx} ${-cy})`
      if (focusEllipse && radial[1]!.toLowerCase() === 'circle' && circleRadius !== undefined) {
        const focusCx = (left + width / 2) * 10
        const focusCy = (top + height / 2) * 10
        const focusRadiusX = width * element.width / 200
        const focusRadiusY = height * element.height / 200
        const focusRadius = (focusRadiusX + focusRadiusY) / 2
        fx = cx + (focusCx - cx) / scaleX
        fy = cy + (focusCy - cy) / scaleY
        fr = focusRadius * 500 / circleRadius
      } else if (focusEllipse) {
        focusOverlay = { kind: 'ellipse', cx: (left + width / 2) * 10, cy: (top + height / 2) * 10, rx: width * 5, ry: height * 5, color: focusColor || '' }
      } else if (focusRect) {
        focusOverlay = { kind: 'rect', x: left * 10, y: top * 10, width: width * 10, height: height * 10, color: focusColor || '' }
      }
    }
    return {
      id: customGradientId(element, kind, mode), kind: 'radial',
      cx, cy, r: 500, fx, fy, fr, colorInterpolation: fr !== undefined ? 'linearRGB' : undefined, gradientTransform, stops: stops as CustomSvgGradientStop[],
      ...(focusOverlay ? { focus: focusOverlay } : {}),
    }
  }

  // DrawingML angles describe the direction of color change; CSS angles already retain that direction.
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill.angle?view=openxml-3.0.1
  // https://www.w3.org/TR/SVG2/pservers/
  const radians = Number(linear![1]) * Math.PI / 180
  const dx = Math.sin(radians), dy = -Math.cos(radians)
  const halfLength = (Math.abs(dx) + Math.abs(dy)) * 500
  return { id: customGradientId(element, kind, mode), kind: 'linear', x1: 500 - dx * halfLength, y1: 500 - dy * halfLength, x2: 500 + dx * halfLength, y2: 500 + dy * halfLength, stops: stops as CustomSvgGradientStop[] }
}

function customGeometryGradients(element: PptxElement): CustomSvgGradient[] {
  const modifiers = [...new Set(element.customPaths?.map(path => path.fillMode).filter(mode => mode))]
  return [
    ...(['fill', 'stroke'] as const).flatMap(kind => {
      const gradient = customSvgGradient(element, kind)
      return gradient ? [gradient] : []
    }),
    ...modifiers.flatMap(mode => { const gradient = customSvgGradient(element, 'fill', mode); return gradient ? [gradient] : [] }),
  ]
}

function customGeometryFillFocus(element: PptxElement, path: NonNullable<PptxElement['customPaths']>[number]) {
  if (!path.fill || !animationOnOff(element, 'fill.on') || animationColor(element, 'fillcolor')) return undefined
  const gradient = customSvgGradient(element, 'fill', path.fillMode)
  const focus = gradient?.kind === 'radial' ? gradient.focus : undefined
  return focus
}

function customGeometryStrokeWidth(element: PptxElement): string {
  return `${animationStrokeWidth(element) * scale.value}px`
}

function customGeometryStrokeStyle(element: PptxElement): CSSProperties {
  // DrawingML custom dash entries are length ratios of the rendered pen width.
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1
  // https://www.w3.org/TR/svg-strokes/
  const strokeDash = animationStrokeDash(element)
  const strokeWidth = animationStrokeWidth(element) * scale.value
  const cap = element.strokeCap || 'butt'
  const pattern = svgStrokeDashPattern(strokeDash, cap)
  // Office presets include round/square caps; custom dashes include round caps only.
  // SVG adds caps; sub-width custom round dashes retain the warned pen-width dot limit.
  // https://github.com/LibreOffice/core/blob/master/oox/source/drawingml/lineproperties.cxx
  return {
    ...(pattern ? { strokeDasharray: pattern.map(length => `${length * strokeWidth}px`).join(' ') } : {}),
    strokeOpacity: !animationColor(element, 'stroke.color') && (/gradient\(/i.test(element.stroke) || Boolean(element.strokePattern)) ? animationOpacity(element, 'stroke.opacity') : undefined,
    strokeLinecap: cap,
  }
}

function cssLineStyle(pattern?: PptxElement['strokeDash']): 'solid' | 'dashed' | 'dotted' {
  if (Array.isArray(pattern)) return 'dashed'
  if (pattern === 'dot' || pattern === 'sysDot') return 'dotted'
  return pattern && pattern !== 'solid' ? 'dashed' : 'solid'
}

function lineStrokeColor(element: PptxElement): string {
  if (!animationOnOff(element, 'stroke.on')) return 'transparent'
  const animatedStroke = animationColor(element, 'stroke.color')
  if (!animatedStroke && element.strokePattern) return `url(#${customGeometryStrokePatternId(element)})`
  return applyColorOpacity(animatedStroke || (element.stroke === 'transparent' ? '#222' : element.stroke), animationOpacity(element, 'stroke.opacity'))
}

function lineRenderStyle(element: PptxElement): CSSProperties {
  const { width, height } = elementRenderSize(element)
  const stroke = animationColor(element, 'stroke.color') || (element.stroke === 'transparent' ? '#222' : element.stroke)
  return { width: `${Math.max(width * scale.value, 1)}px`, height: `${Math.max(height * scale.value, 1)}px`, color: applyColorOpacity(stroke, animationOpacity(element, 'stroke.opacity')) }
}

function lineEndId(element: PptxElement, side: 'head' | 'tail'): string {
  return `pptx-line-${element.id}-${side}`
}

function lineEndUrl(element: PptxElement, side: 'head' | 'tail'): string | undefined {
  const ending = animationLineEnd(element, side)
  return ending && ending.type !== 'none' ? `url(#${lineEndId(element, side)})` : undefined
}

function lineEndMarkerSize(element: PptxElement, side: 'head' | 'tail'): { width: number; height: number } {
  const ending = animationLineEnd(element, side)
  const size: Record<NonNullable<typeof ending>['width'], number> = { sm: 0.7, med: 1, lg: 1.4 }
  return { width: 6 * size[ending.length], height: 6 * size[ending.width] }
}

function lineEndPath(type: NonNullable<PptxElement['lineHeadEnd']>['type']): string {
  if (type === 'stealth') return 'M0 0L10 5L0 10L3.5 5Z'
  if (type === 'diamond') return 'M0 5L5 0L10 5L5 10Z'
  if (type === 'oval') return 'M5 0A5 5 0 1 1 5 10A5 5 0 1 1 5 0Z'
  if (type === 'arrow' || type === 'chevron') return 'M0 0L10 5L0 10'
  if (type === 'doubleChevron') return 'M0 0L6 5L0 10M4 0L10 5L4 10'
  return 'M0 0L10 5L0 10Z'
}

function lineEndUsesStroke(type: NonNullable<PptxElement['lineHeadEnd']>['type']): boolean {
  return type === 'arrow' || type === 'chevron' || type === 'doubleChevron'
}

function lineArrowType(value: NonNullable<PptxAnimation['lineArrowFrom']>): NonNullable<PptxElement['lineHeadEnd']>['type'] {
  if (value === 'block') return 'triangle'
  if (value === 'classic') return 'stealth'
  if (value === 'open') return 'arrow'
  if (value === 'doublechevron') return 'doubleChevron'
  return value
}

function lineArrowSize(value: NonNullable<PptxAnimation['lineArrowSizeFrom']>): NonNullable<PptxElement['lineHeadEnd']>['width'] {
  if (value === 'narrow' || value === 'short') return 'sm'
  if (value === 'wide' || value === 'long') return 'lg'
  return 'med'
}

function animationLineEnd(element: PptxElement, side: 'head' | 'tail'): NonNullable<PptxElement['lineHeadEnd']> {
  const property = side === 'head' ? 'stroke.startArrow' : 'stroke.endArrow'
  const baseline = side === 'head' ? element.lineHeadEnd : element.lineTailEnd
  let type = baseline?.type || 'none'
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'lineArrow' && item.lineArrowProperty === property) || []) {
    const progress = animationProgress(action)
    if (progress === -2) {
      type = baseline?.type || 'none'
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.lineArrowKeyframes?.length) {
      for (const frame of action.lineArrowKeyframes) {
        if (frame.offset > progress) break
        type = lineArrowType(frame.value)
      }
    } else if (action.lineArrowFrom !== undefined && progress < 1) {
      type = lineArrowType(action.lineArrowFrom)
    } else if (progress >= 1) {
      const value = action.lineArrowTo ?? action.lineArrowFrom
      if (value !== undefined) type = lineArrowType(value)
    }
  }
  const baselineWidth = baseline?.width || 'med'
  const baselineLength = baseline?.length || 'med'
  const animatedSize = (property: NonNullable<PptxAnimation['lineArrowSizeProperty']>, initial: typeof baselineWidth): typeof baselineWidth => {
    let value = initial
    for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'lineArrowSize' && item.lineArrowSizeProperty === property) || []) {
      const progress = animationProgress(action)
      if (progress === -2) {
        value = initial
        continue
      }
      if (progress === null || progress < 0) continue
      if (action.lineArrowSizeKeyframes?.length) {
        for (const frame of action.lineArrowSizeKeyframes) {
          if (frame.offset > progress) break
          value = lineArrowSize(frame.value)
        }
      } else if (action.lineArrowSizeFrom !== undefined && progress < 1) {
        value = lineArrowSize(action.lineArrowSizeFrom)
      } else if (progress >= 1) {
        const finalValue = action.lineArrowSizeTo ?? action.lineArrowSizeFrom
        if (finalValue !== undefined) value = lineArrowSize(finalValue)
      }
    }
    return value
  }
  const widthProperty = side === 'head' ? 'stroke.startArrowWidth' : 'stroke.endArrowWidth'
  const lengthProperty = side === 'head' ? 'stroke.startArrowLength' : 'stroke.endArrowLength'
  return {
    type,
    width: animatedSize(widthProperty, baselineWidth),
    length: animatedSize(lengthProperty, baselineLength)
  }
}

function customGeometryStyle(element: PptxElement): CSSProperties {
  const shadows = element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []
  const outerShadows = shadows.filter(shadow => !shadow.inset)
  const { scaleX, scaleY } = elementRenderSize(element)
  const groupScale = { scaleX, scaleY }
  const innerShadowFilter = customGeometryInnerShadowFilterReference(element)
  if ((!outerShadows.length && !innerShadowFilter) || !animationOnOff(element, 'shadow.on')) return {}
  return {
    filter: [
      innerShadowFilter,
      ...outerShadows.map(shadow => `drop-shadow(${shadowCss(shadow, applyColorOpacity(animationColor(element, 'shadow.color') || shadow.color, animationOpacity(element, 'shadow.opacity')), 0.5, groupScale, animationShadowOffset(element, shadow, shadows.indexOf(shadow)))})`)
    ].filter((filter): filter is string => Boolean(filter)).join(' ')
  }
}

function customGeometryInnerShadows(element: PptxElement): PptxShadow[] {
  const shadows = element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []
  return shadows.filter(shadow => shadow.inset)
}

function customGeometryInnerShadowFilterId(element: PptxElement): string {
  return `${customGradientScope}-inner-shadow-${element.id}`.replace(/[^a-zA-Z0-9_-]/g, '_')
}

function customGeometryInnerShadowFilterReference(element: PptxElement): string | undefined {
  return customGeometryInnerShadows(element).length && animationOnOff(element, 'shadow.on')
    ? `url(#${customGeometryInnerShadowFilterId(element)})`
    : undefined
}

function customGeometryInnerShadowFilterRegion(element: PptxElement) {
  let paddingX = 0
  let paddingY = 0
  for (const shadow of customGeometryInnerShadows(element)) {
    const angle = (shadow.angle || 0) * Math.PI / 180
    const blur = Math.max(0, shadow.blur) * 1.5
    paddingX = Math.max(paddingX, (Math.abs(Math.cos(angle) * shadow.distance) + blur) / Math.max(1, element.width))
    paddingY = Math.max(paddingY, (Math.abs(Math.sin(angle) * shadow.distance) + blur) / Math.max(1, element.height))
  }
  return {x: -paddingX, y: -paddingY, width: 1 + paddingX * 2, height: 1 + paddingY * 2}
}

function customGeometryInnerShadowDeviation(element: PptxElement, shadow: PptxShadow): string {
  const blur = Math.max(0, shadow.blur) / 2
  return `${blur / Math.max(1, element.width)} ${blur / Math.max(1, element.height)}`
}

function customGeometryInnerShadowOffset(element: PptxElement, shadow: PptxShadow): {dx: number; dy: number} {
  const [x, y] = animationShadowOffset(element, shadow, Math.max(0, element.shadows?.indexOf(shadow) ?? 0))
  return {
    dx: x / Math.max(1, element.width),
    dy: y / Math.max(1, element.height)
  }
}

function customGeometryInnerShadowColor(element: PptxElement, shadow: PptxShadow): string {
  return applyColorOpacity(animationColor(element, 'shadow.color') || shadow.color, animationOpacity(element, 'shadow.opacity'))
}

function animationShadowOffset(element: PptxElement, shadow: PptxShadow, shadowIndex = 0, paragraphIndex?: number, characterStart?: number, characterEnd?: number): [number, number] {
  const angle = (shadow.angle || 0) * Math.PI / 180
  const baseX = Math.cos(angle) * shadow.distance
  const baseY = Math.sin(angle) * shadow.distance
  let x = baseX
  let y = baseY
  if (shadowIndex !== 0) return [x, y]
  const actions = animationsByTarget.value.get(element.id)?.filter(item => {
    if (item.effect !== 'shadowOffset') return false
    const hasTextTarget = Boolean(item.paragraphRange || item.characterRange || item.iteration)
    if (!hasTextTarget) return true
    if (paragraphIndex === undefined) return false
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(item, characterStart, characterEnd)
    return (!item.paragraphRange || paragraphIndex >= item.paragraphRange.start && paragraphIndex <= item.paragraphRange.end)
      && (!item.characterRange || characterStart !== undefined && characterEnd !== undefined && characterStart < item.characterRange.end && characterEnd > item.characterRange.start)
      && (!item.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  for (const action of actions) {
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(action, characterStart, characterEnd)
    const progress = animationProgress(action, iterationIndex)
    if (progress === -2) {
      if (action.shadowOffsetProperty === 'shadow.offset.x') x = baseX
      else if (action.shadowOffsetProperty === 'shadow.offset.y') y = baseY
      continue
    }
    if (progress === null || progress < 0 || action.shadowOffsetFrom === undefined || action.shadowOffsetTo === undefined) continue
    const keyframes = action.shadowOffsetKeyframes
    const formulaSamples = action.shadowOffsetFormulaSamples
    const from = action.shadowOffsetFrom
    const to = action.shadowOffsetTo
    let value = formulaSamples?.length
      ? animationKeyframeValue(formulaSamples, 'lin', progress)
      : keyframes?.length
        ? animationKeyframeValue(keyframes, action.shadowOffsetKeyframeMode, progress)
        : from + (to - from) * progress
    if (action.shadowOffsetAccumulate) {
      value += (to - from) * completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1)
    }
    if (action.shadowOffsetProperty === 'shadow.offset.x') x = value
    else if (action.shadowOffsetProperty === 'shadow.offset.y') y = value
  }
  return [x, y]
}

function shadowCss(shadow: PptxShadow, color: string, blurScale = 1, groupScale = { scaleX: 1, scaleY: 1 }, offset?: [number, number]): string {
  const angle = (shadow.angle || 0) * Math.PI / 180
  const blurTransformScale = Math.sqrt(groupScale.scaleX * groupScale.scaleY)
  const [offsetX, offsetY] = offset || [Math.cos(angle) * shadow.distance, Math.sin(angle) * shadow.distance]
  return [
    ...(shadow.inset ? ['inset'] : []),
    offsetX * scale.value * groupScale.scaleX + 'px',
    offsetY * scale.value * groupScale.scaleY + 'px',
    shadow.blur * blurScale * scale.value * blurTransformScale + 'px',
    color
  ].join(' ')
}

function pictureInnerShadowStyle(element: PptxElement): CSSProperties | undefined {
  const shadows = element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []
  const innerShadows = shadows.filter(shadow => shadow.inset)
  if (!innerShadows.length || !animationOnOff(element, 'shadow.on')) return undefined
  const { scaleX, scaleY } = elementRenderSize(element)
  return {
    boxShadow: innerShadows.map(shadow => shadowCss(shadow, applyColorOpacity(animationColor(element, 'shadow.color') || shadow.color, animationOpacity(element, 'shadow.opacity')), 1, { scaleX, scaleY }, animationShadowOffset(element, shadow, shadows.indexOf(shadow)))).join(', ')
  }
}

function imageStyle(element: PptxElement): CSSProperties {
  const hasCropAnimation = animationsByTarget.value.get(element.id)?.some(item => item.effect === 'imageCrop')
  const crop = element.imageCrop || hasCropAnimation ? imageCrop(element) : undefined
  const grayscale = element.kind === 'picture' && imageGrayscale(element)
  const alphaInverse = element.kind === 'picture' && element.imageAlphaInverse !== undefined ? `url(#${imageAlphaInverseFilterId(element)})` : undefined
  const effectDagAlphaInverse = element.kind === 'picture' && element.imageEffectDagAlphaInverse !== undefined ? `url(#${imageAlphaInverseFilterId(element, true)})` : undefined
  const grayscaleFilter = grayscale ? 'grayscale(1)' : undefined
  const imageEffects = element.imageAlphaInverseAfterGrayscale ? [grayscaleFilter, alphaInverse] : [alphaInverse, grayscaleFilter]
  const filters = [
    imageToneEnabled(element) ? `url(#${imageToneFilterId(element)})` : undefined,
    ...imageEffects,
    effectDagAlphaInverse,
  ].filter((filter): filter is string => Boolean(filter))
  const filter = filters.length ? filters.join(' ') : undefined
  if (!crop) return { width: '100%', height: '100%', filter }
  const width = 1 - crop.left - crop.right
  const height = 1 - crop.top - crop.bottom
  return {
    width: `${100 / width}%`,
    height: `${100 / height}%`,
    left: `${-crop.left / width * 100}%`,
    top: `${-crop.top / height * 100}%`,
    filter,
  }
}

function imageCrop(element: PptxElement): NonNullable<PptxElement['imageCrop']> {
  const base = element.imageCrop || {left: 0, top: 0, right: 0, bottom: 0}
  const crop = {...base}
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'imageCrop') || []) {
    const property = action.imageCropProperty
    if (!property || (action.imageCropTo === undefined && !action.imageCropKeyframes?.length)) continue
    const progress = animationProgress(action)
    if (progress === -2) {
      crop[property] = base[property]
      continue
    }
    if (progress === null || progress < 0) continue
    const value = action.imageCropKeyframes?.length
      ? animationKeyframeValue(action.imageCropKeyframes, action.imageCropKeyframeMode, progress)
      : (action.imageCropFrom ?? crop[property]) + ((action.imageCropTo ?? 0) - (action.imageCropFrom ?? crop[property])) * progress
    const opposite = property === 'left' ? 'right' : property === 'right' ? 'left' : property === 'top' ? 'bottom' : 'top'
    if (value >= 0 && value < 1 && value + crop[opposite] < 1) crop[property] = value
  }
  return crop
}

function imageGrayscale(element: PptxElement): boolean {
  let grayscale = element.imageGrayscale === true
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'grayscale') || []) {
    const progress = animationProgress(action)
    if (progress === null || progress < 0) continue
    let value: boolean = action.grayscaleFrom ?? grayscale
    if (action.grayscaleKeyframes?.length) {
      for (const frame of action.grayscaleKeyframes) {
        if (frame.offset > progress) break
        value = frame.value
      }
    } else if (progress >= 1) value = action.grayscaleTo ?? value
    grayscale = value
  }
  return grayscale
}

function imageTone(element: PptxElement): {gain: number; blacklevel: number; gamma: number} {
  const tone = {gain: 1, blacklevel: 0, gamma: 1}
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'imageTone') || []) {
    const property = action.imageToneProperty
    if (!property || (action.imageToneTo === undefined && !action.imageToneKeyframes?.length)) continue
    const progress = animationProgress(action)
    if (progress === -2) {
      tone[property] = property === 'blacklevel' ? 0 : 1
      continue
    }
    if (progress === null || progress < 0) continue
    const from = action.imageToneFrom ?? tone[property]
    const value = action.imageToneKeyframes?.length
      ? animationKeyframeValue(action.imageToneKeyframes, action.imageToneKeyframeMode, progress)
      : from + ((action.imageToneTo ?? from) - from) * progress
    if (Number.isFinite(value)) tone[property] = value
  }
  return tone
}

function imageToneEnabled(element: PptxElement): boolean {
  return element.kind === 'picture' && Boolean(animationsByTarget.value.get(element.id)?.some(item => item.effect === 'imageTone'))
}

function imageToneFilterId(element: PptxElement): string {
  return `${customGradientScope}-image-tone-${element.id}`.replace(/[^a-zA-Z0-9_-]/g, '_')
}

function imageAlphaInverseFilterId(element: PptxElement, effectDag = false): string {
  return (customGradientScope + '-image-alpha-inverse-' + element.id + (effectDag ? '-effect-dag' : '')).replace(/[^a-zA-Z0-9_-]/g, '_')
}

function mediaPlaybackRange(media: HTMLMediaElement, element: PptxElement) {
  const trim = element.mediaTrim
  if (!trim) return { start: 0, end: media.duration, valid: true }
  const start = trim.startMs / 1000
  const end = media.duration - trim.endMs / 1000
  if (start < end) return { start, end, valid: true }
  if (!invalidMediaTrimIds.value.includes(element.id)) invalidMediaTrimIds.value = [...invalidMediaTrimIds.value, element.id]
  return { start: 0, end: media.duration, valid: false }
}

function applyMediaTrim(event: Event, element: PptxElement) {
  const media = event.currentTarget as HTMLMediaElement
  if (!element.mediaTrim || media.readyState === 0 || !Number.isFinite(media.duration)) return
  const { start, end, valid } = mediaPlaybackRange(media, element)
  if (!valid) return
  if (media.currentTime < start) media.currentTime = start
  else if (media.currentTime >= end) {
    media.pause()
    if (media.currentTime > end + 0.01) media.currentTime = end
  }
}

function mediaFadeFactor(media: HTMLMediaElement, element: PptxElement): number {
  const fade = element.mediaFade
  if (!fade) return 1
  const { start, end } = mediaPlaybackRange(media, element)
  if (fade.inMs + fade.outMs > (end - start) * 1_000) {
    if (!invalidMediaFadeIds.value.includes(element.id)) invalidMediaFadeIds.value = [...invalidMediaFadeIds.value, element.id]
    return 1
  }
  let factor = 1
  if (fade.inMs) factor = Math.min(factor, (media.currentTime - start) / (fade.inMs / 1_000))
  if (fade.outMs) factor = Math.min(factor, (end - media.currentTime) / (fade.outMs / 1_000))
  return Math.max(0, Math.min(factor, 1))
}

function updateMediaFade(media: HTMLMediaElement, element: PptxElement) {
  if (element.mediaFade && Number.isFinite(media.duration)) {
    const factor = mediaFadeFactor(media, element)
    let state = mediaFadeStates.get(media)
    if (!state) {
      state = { baseVolume: media.volume, appliedVolume: media.volume, appliedFactor: 1, baseMuted: media.muted, appliedMuted: media.muted }
      mediaFadeStates.set(media, state)
    } else {
      if (Math.abs(media.volume - state.appliedVolume) > 0.005) {
        state.baseVolume = Math.max(0, Math.min(state.appliedFactor ? media.volume / state.appliedFactor : media.volume, 1))
      }
      if (media.muted !== state.appliedMuted) state.baseMuted = media.muted
    }
    const appliedVolume = state.baseVolume * (factor > 0 ? factor : 1)
    const appliedMuted = state.baseMuted || factor === 0
    state.appliedVolume = appliedVolume
    state.appliedFactor = factor
    state.appliedMuted = appliedMuted
    if (Math.abs(media.volume - appliedVolume) > 0.001) media.volume = appliedVolume
    if (media.muted !== appliedMuted) media.muted = appliedMuted
  }
  const currentFrame = mediaPlaybackFrames.get(media)
  const hasBookmarkTriggers = mediaBookmarkTriggersByMedia.value.has(element.id)
  if (media.paused || media.ended || !media.isConnected || (!element.mediaFade && !hasBookmarkTriggers)) {
    if (currentFrame !== undefined) cancelAnimationFrame(currentFrame)
    mediaPlaybackFrames.delete(media)
  } else if (currentFrame === undefined) {
    mediaPlaybackFrames.set(media, requestAnimationFrame(() => {
      mediaPlaybackFrames.delete(media)
      if (!media.paused && !media.ended && media.isConnected && mediaBookmarkTriggersByMedia.value.has(element.id)) {
        runMediaBookmarkAnimations(media, element, !media.seeking)
      }
      updateMediaFade(media, element)
    }))
  }
}

function updateMediaPlayback(event: Event, element: PptxElement) {
  const media = event.currentTarget as HTMLMediaElement
  const previousTime = lastMediaBookmarkTimes.get(media)
  if (event.type === 'seeking' && previousTime !== undefined && media.currentTime < previousTime) mediaBookmarkPlaybackStarted.delete(media)
  applyMediaTrim(event, element)
  updateMediaFade(media, element)
  runMediaBookmarkAnimations(media, element, (event.type === 'timeupdate' || event.type === 'play') && !media.seeking)
}

function updateMediaFadeForEvent(event: Event, element: PptxElement) {
  updateMediaFade(event.currentTarget as HTMLMediaElement, element)
}

function updateMediaOnPlay(event: Event, element: PptxElement) {
  blockedMediaPlaybackIds.value = blockedMediaPlaybackIds.value.filter(id => id !== element.id)
  let startedAction = false
  timelineActions.value.forEach((action, index) => {
    if (action.effect !== 'media' || !['play', 'togglePause'].includes(action.mediaCommand || '') || action.targetId !== element.id || !action.mediaActionKey || !executedMediaActions.has(index) || startedMediaActions.has(action.mediaActionKey)) return
    startedMediaActions.set(action.mediaActionKey, performance.now())
    startedAction = true
  })
  syncExclusiveMediaPlayback(performance.now())
  updateMediaPlayback(event, element)
  if (startedAction) refreshTimeline()
}

function finishTimedMedia(target: PptxElement | string, refresh = true, targetEnded = false) {
  const targetId = typeof target === 'string' ? target : target.id
  const endedAt = performance.now()
  let changed = false
  const media = Array.from(stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || []).find(item => item.dataset.mediaId === targetId)
  if (media) {
    delete media.dataset.mediaActionEndAt
    mediaBookmarkPlaybackStarted.delete(media)
    lastMediaBookmarkTimes.delete(media)
  }
  timelineActions.value.forEach((action, index) => {
    if (action.effect !== 'media' || action.targetId !== targetId || !action.mediaActionKey || !executedMediaActions.has(index)) return
    completedMediaActions.set(action.mediaActionKey, endedAt)
    if (targetEnded) endedMediaTargets.set(action.mediaActionKey, endedAt)
    clearMediaActionEnd(action.mediaActionKey)
    changed = true
  })
  if (changed && refresh) refreshTimeline()
}

function runTimedMediaActions() {
  timelineActions.value.forEach((action, index) => {
    if (action.effect !== 'media' || !action.mediaCommand || executedMediaActions.has(index)) return
    const startedAt = timelineStart(action)
    const stoppedAt = timelineEnd(action)
    const actionNow = timelineNow(action)
    if (stoppedAt !== undefined && startedAt !== undefined && startedAt + action.delayMs >= timelineTimeAt(action, stoppedAt)) return
    if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY || timelineIsPaused(action, animationClock.value) || actionNow < startedAt + action.delayMs) return
    executedMediaActions.add(index)
    if (action.mediaActionKey) scheduledMediaActions.set(action.mediaActionKey, animationClock.value)
    if (action.mediaCommand === 'stopAudio') {
      emit('stop-all-audio')
      for (const element of props.slide.elements) {
        if (element.kind !== 'audio') continue
        const audio = Array.from(stage.value?.querySelectorAll<HTMLAudioElement>('[data-media-id]') || []).find(item => item.dataset.mediaId === element.id)
        if (!audio) continue
        if (audio.readyState > 0) audio.currentTime = (element.mediaTrim?.startMs || 0) / 1000
        finishTimedMedia(element, false)
      }
      for (const audio of stage.value?.querySelectorAll<HTMLAudioElement>('audio[data-animation-sound]') || []) {
        audio.pause()
        if (audio.readyState > 0) audio.currentTime = 0
        if (audio.dataset.mediaId) finishTimedMedia(audio.dataset.mediaId, false)
      }
      runSlideTriggeredAnimations('onStopAudio')
      refreshTimeline()
      return
    }
    const element = props.slide.elements.find(item => item.id === action.targetId)
    const media = Array.from(stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || []).find(item => item.dataset.mediaId === action.targetId)
    if (!media || (!element && !action.mediaSourceUrl)) return
    if (action.mediaSourceUrl) {
      if (action.mediaCommand !== 'play') return
      if (action.mediaVolume !== undefined) media.volume = action.mediaVolume
      if (action.mediaMuted !== undefined) media.muted = action.mediaMuted
      media.currentTime = 0
      void media.play().then(() => {
        if (action.mediaActionKey && !startedMediaActions.has(action.mediaActionKey)) {
          startedMediaActions.set(action.mediaActionKey, performance.now())
          refreshTimeline()
        }
      }).catch(error => {
        if (error instanceof DOMException && error.name === 'NotAllowedError' && !blockedMediaPlaybackIds.value.includes(action.targetId)) {
          blockedMediaPlaybackIds.value = [...blockedMediaPlaybackIds.value, action.targetId]
        }
      })
      if (action.mediaActionKey && action.mediaDurationMs !== undefined) {
        scheduleMediaActionEnd(action, media, action.mediaDurationMs)
      }
      return
    }
    if (!element) return
    const startsPlayback = action.mediaCommand === 'play' || action.mediaCommand === 'togglePause' && media.paused
    if (action.mediaActionKey && (!startsPlayback || !media.paused)) startedMediaActions.set(action.mediaActionKey, performance.now())
    if (action.mediaSlideCount) media.dataset.mediaSlideCount = String(action.mediaSlideCount)
    if (action.mediaVolume !== undefined) media.volume = action.mediaVolume
    if (action.mediaMuted !== undefined) media.muted = action.mediaMuted
    const seek = (seconds: number) => {
      const apply = () => {
        if (media.readyState === 0 || !Number.isFinite(media.duration)) return
        const range = mediaPlaybackRange(media, element)
        media.currentTime = range.valid ? Math.max(range.start, Math.min(seconds, range.end)) : Math.max(0, Math.min(seconds, media.duration))
      }
      if (media.readyState === 0 || !Number.isFinite(media.duration)) media.addEventListener('loadedmetadata', apply, { once: true })
      else apply()
    }
    const play = () => {
      void media.play().then(() => {
        if (action.mediaActionKey && !startedMediaActions.has(action.mediaActionKey)) {
          startedMediaActions.set(action.mediaActionKey, performance.now())
          refreshTimeline()
        }
      }).catch(error => {
        if (error instanceof DOMException && error.name === 'NotAllowedError' && !blockedMediaPlaybackIds.value.includes(element.id)) {
          blockedMediaPlaybackIds.value = [...blockedMediaPlaybackIds.value, element.id]
        }
      })
    }
    if (action.mediaCommand === 'play') {
      if (action.mediaStartSeconds !== undefined) seek(action.mediaStartSeconds)
      else if (media.currentTime < (element.mediaTrim?.startMs || 0) / 1000) seek((element.mediaTrim?.startMs || 0) / 1000)
      play()
      if (action.mediaActionKey && action.mediaDurationMs !== undefined) {
        scheduleMediaActionEnd(action, media, action.mediaDurationMs)
      }
    } else if (action.mediaCommand === 'pause') media.pause()
    else if (action.mediaCommand === 'togglePause') {
      if (media.paused) play()
      else media.pause()
    } else {
      media.pause()
      seek((element.mediaTrim?.startMs || 0) / 1000)
      finishTimedMedia(element, false)
    }
  })
}

function easedProgress(action: PptxAnimation, progress: number): number {
  return timingProgress(action, progress)
}

function finalAnimationProgress(action: PptxAnimation): number {
  if (action.effectProgress !== undefined) return action.effectProgress
  if (action.durationMs === 0) return action.reversePlayback ? 0 : 1
  const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1)
  const totalDuration = animationDuration(action)
  if (totalDuration === 0) return 0
  if (!Number.isFinite(totalDuration)) return 0
  const remainder = totalDuration % cycleDuration
  let progress: number
  if (!action.autoReverse) progress = easedProgress(action, remainder === 0 ? 1 : remainder / action.durationMs)
  else if (remainder === 0) progress = 0
  else if (remainder <= action.durationMs) progress = easedProgress(action, remainder / action.durationMs)
  else progress = 1 - easedProgress(action, (remainder - action.durationMs) / action.durationMs)
  return action.reversePlayback ? 1 - progress : progress
}

function animationFillExpired(action: RenderAnimation, startedAt: number | undefined, iterationIndex = 0): boolean {
  if (action.fillMode === 'remove') {
    if (startedAt === undefined) return action.triggerSequence === undefined && !action.automatic && action.step < props.animationStep
    if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return true
    const stoppedAt = timelineEnd(action)
    if (stoppedAt !== undefined && animationClock.value >= stoppedAt) return true
    const stagger = action.iteration ? action.iteration.intervalMs * iterationIndex : 0
    return action.timingWarp
      ? timingWarpChildElapsed(action, startedAt, stagger) >= animationDuration(action)
      : timelineNow(action) >= startedAt + action.delayMs + stagger + animationDuration(action)
  }
  if (action.fillMode !== 'freeze' || startedAt === undefined) return false
  const actionStart = startedAt + action.delayMs
  // Cross-target actions expire fills at later click steps or nested sequence children.
  return timelineActions.value.some((sibling) => {
    if (!animationStartsLaterInFillScope(action, sibling)) return false
    const siblingStart = timelineStart(sibling)
    if (siblingStart === undefined) return false
    const start = siblingStart + sibling.delayMs
    return start >= actionStart && start <= timelineNow(action)
  })
}

function animationStartsLaterInFillScope(action: RenderAnimation, sibling: RenderAnimation): boolean {
  const actionScopes = action.fillSequenceScopes || (action.sequenceKey === undefined ? [] : [{key: action.sequenceKey, order: action.sequenceOrder ?? -1}])
  const siblingScopes = sibling.fillSequenceScopes || (sibling.sequenceKey === undefined ? [] : [{key: sibling.sequenceKey, order: sibling.sequenceOrder ?? -1}])
  return actionScopes.some(scope => siblingScopes.some(siblingScope => {
    if (siblingScope.key !== scope.key || siblingScope.order <= scope.order) return false
    if (scope.key !== action.sequenceKey) return true
    if (action.triggerSequence !== undefined) {
      return sibling.triggerSequence === action.triggerSequence && (sibling.triggerStep ?? -1) > (action.triggerStep ?? -1)
    }
    return sibling.step > action.step
  }))
}

function timingWarpChildElapsed(action: RenderAnimation, startedAt: number, stagger = 0): number {
  const warp = action.timingWarp!
  const iterationStartOffsetMs = warp.groupPerIteration ? stagger : 0
  const elapsedWallMs = timelineNow(action) - startedAt - warp.groupOffsetMs - iterationStartOffsetMs
  if (elapsedWallMs < 0) return -1
  const elapsedPlaybackMs = Math.min(elapsedWallMs, warp.groupPlaybackDurationMs) * warp.groupSpeed
  const cycleDurationMs = warp.groupDurationMs * (warp.groupAutoReverse ? 2 : 1)
  let cyclePositionMs = elapsedPlaybackMs % cycleDurationMs
  const atEnd = elapsedWallMs >= warp.groupPlaybackDurationMs
  if (atEnd && (cyclePositionMs < 0.01 || cycleDurationMs - cyclePositionMs < 0.01)) cyclePositionMs = cycleDurationMs
  const reverseLeg = warp.groupAutoReverse && cyclePositionMs > warp.groupDurationMs
  const legProgress = reverseLeg
    ? (cyclePositionMs - warp.groupDurationMs) / warp.groupDurationMs
    : cyclePositionMs / warp.groupDurationMs
  const isReverse = warp.groupReversePlayback !== reverseLeg
  const groupProgress = timingProgress(warp, isReverse ? 1 - legProgress : legProgress)
  return groupProgress * warp.groupDurationMs
    - warp.childOffsetMs - (warp.groupPerIteration ? 0 : stagger)
}

function animationProgress(action: RenderAnimation, iterationIndex?: number): number | null {
  if (props.thumbnail && !props.playbackPreview) return null
  const now = timelineNow(action)
  if (action.iteration && (iterationIndex === undefined || iterationIndex < 0)) return null
  const startedAt = timelineStart(action)
  const stagger = action.iteration ? action.iteration.intervalMs * (iterationIndex || 0) : 0
  if (animationFillExpired(action, startedAt, iterationIndex || 0)) return -2
  if (startedAt === undefined) return action.triggerSequence === undefined && !action.automatic && action.step < props.animationStep ? finalAnimationProgress(action) : null
  if (startedAt === Infinity) return -1
  if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return finalAnimationProgress(action)
  const elapsed = action.timingWarp
    ? timingWarpChildElapsed(action, startedAt, stagger)
    : now - startedAt - action.delayMs - stagger
  if (elapsed < 0) return -1
  if (action.effectProgress !== undefined) return action.effectProgress
  if (action.durationMs === 0) return action.reversePlayback ? 0 : 1
  const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1)
  const totalDuration = animationDuration(action)
  if (totalDuration === 0) return 0
  if (elapsed >= totalDuration) return finalAnimationProgress(action)
  const position = elapsed % cycleDuration
  const progress = action.autoReverse && position > action.durationMs
    ? 1 - easedProgress(action, (position - action.durationMs) / action.durationMs)
    : easedProgress(action, Math.min(position / action.durationMs, 1))
  return action.reversePlayback ? 1 - progress : progress
}

function completedAnimationRepeats(action: RenderAnimation, iterationIndex?: number): number {
  const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1)
  if (!(action.rotationAccumulate || action.scaleAccumulate || action.motionAccumulate || action.opacityAccumulate || action.shadowOffsetAccumulate) || action.autoReverse || cycleDuration <= 0) return 0
  const totalDuration = animationDuration(action)
  const lastRepeat = Number.isFinite(totalDuration) ? Math.max(0, Math.ceil(totalDuration / cycleDuration) - 1) : Infinity
  const startedAt = timelineStart(action)
  if (startedAt === undefined) return action.triggerSequence === undefined && !action.automatic && action.step < props.animationStep ? lastRepeat : 0
  if (startedAt === Infinity) return 0
  if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return Number.isFinite(lastRepeat) ? lastRepeat : 0
  const stagger = action.iteration ? action.iteration.intervalMs * (iterationIndex || 0) : 0
  const elapsed = action.timingWarp
    ? timingWarpChildElapsed(action, startedAt, stagger)
    : timelineNow(action) - startedAt - action.delayMs - stagger
  return Math.min(lastRepeat, Math.max(0, Math.floor(elapsed / cycleDuration)))
}

function motionAccumulationOffset(action: RenderAnimation, iterationIndex?: number): [number, number] {
  if (!action.motionAccumulate) return [0, 0]
  const points = action.motionPathSamples || action.motionPath
  const start = points?.[0]
  const end = points?.[points.length - 1]
  if (!start || !end) return [0, 0]
  const repeats = completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1)
  return [(end.x - start.x) * repeats, (end.y - start.y) * repeats]
}

type AnimationScaleLayers = { dimensions: [number, number]; transform: [number, number] }

function applyAnimationScaleLayer(layers: AnimationScaleLayers, action: RenderAnimation, value: [number, number]) {
  const layer = action.scaleLayer === 'dimensions' ? layers.dimensions : layers.transform
  const applyValue = (axis: 0 | 1) => {
    layer[axis] = action.scaleBaseAdditive
      ? 1 + value[axis]
      : action.scaleAdditive
      ? layer[axis] + value[axis]
      : action.scaleMultiplicative
        ? layer[axis] * value[axis]
        : value[axis]
  }
  if (action.scaleAxis === 'x') applyValue(0)
  else if (action.scaleAxis === 'y') applyValue(1)
  else {
    applyValue(0)
    applyValue(1)
  }
}

function composedAnimationScale(layers: AnimationScaleLayers): [number, number] {
  return [layers.dimensions[0] * layers.transform[0], layers.dimensions[1] * layers.transform[1]]
}

function animationScale(element: PptxElement): [number, number] {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => ['scale', 'zoom'].includes(action.effect) && !action.paragraphRange && !action.characterRange && !action.iteration) || []
  const layers: AnimationScaleLayers = { dimensions: [1, 1], transform: [1, 1] }
  for (const action of actions) {
    const progress = animationProgress(action)
    if (progress === -2) {
      if (!action.scaleAdditive && !action.scaleMultiplicative) applyAnimationScaleLayer(layers, action, [1, 1])
      continue
    }
    if (progress === null || progress < 0) {
      if (action.effect === 'zoom' && action.direction === 'in' && action.scaleFrom) applyAnimationScaleLayer(layers, action, action.scaleFrom)
      continue
    }
    if (!action.scaleFrom || !action.scaleTo) continue
    const currentScale = action.scaleLayer === 'dimensions' ? layers.dimensions : layers.transform
    const animatedScale = animationScaleValue(action, progress, undefined, action.scaleFromCurrent ? currentScale : undefined)
    applyAnimationScaleLayer(layers, action, animatedScale)
  }
  return composedAnimationScale(layers)
}

function animationScaleValue(action: RenderAnimation, progress: number, iterationIndex?: number, currentScale?: [number, number]): [number, number] {
  const [fromX, fromY] = currentScale && action.scaleFromCurrent ? currentScale : action.scaleFrom!
  const [toX, toY] = action.scaleTo!
  const completedRepeats = action.scaleAccumulate
    ? completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1)
    : 0
  const valueAt = (axis: 'x' | 'y', from: number, to: number) => {
    const frames = action.scaleKeyframes
    if (!frames?.length || action.scaleAxis !== axis) return from + (to - from) * progress
    let low = 0
    let high = frames.length
    while (low < high) {
      const middle = (low + high) >>> 1
      if (frames[middle]!.offset < progress) low = middle + 1
      else high = middle
    }
    if (low === 0) return frames[0]!.value
    if (low === frames.length) {
      if (action.scaleKeyframeMode === 'discrete' || frames.length < 2) return frames.at(-1)!.value
      const before = frames.at(-2)!, after = frames.at(-1)!
      return before.value + (after.value - before.value) * (progress - before.offset) / (after.offset - before.offset)
    }
    const before = frames[low - 1]!
    const after = frames[low]!
    if (action.scaleKeyframeMode === 'discrete') return progress === after.offset ? after.value : before.value
    const ratio = (progress - before.offset) / (after.offset - before.offset)
    return before.value + (after.value - before.value) * ratio
  }
  return [
    valueAt('x', fromX, toX) + (toX - fromX) * completedRepeats,
    valueAt('y', fromY, toY) + (toY - fromY) * completedRepeats,
  ]
}

function animationRotation(element: PptxElement): number {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'rotation' && !action.paragraphRange && !action.characterRange && !action.iteration) || []
  let value = element.rotation
  for (const action of actions) {
    const progress = animationProgress(action)
    if (progress === -2) {
      value = element.rotation
      continue
    }
    if (progress === null || progress < 0 || action.rotationFrom === undefined || action.rotationTo === undefined) continue
    const rotation = animationRotationValue(action, progress, value)
    value = action.rotationRelative || action.rotationAdditive ? value + rotation : rotation
  }
  return value
}

function animationRotationValue(action: RenderAnimation, progress: number, currentRotation: number, iterationIndex?: number): number {
  const from = action.rotationFromCurrent ? currentRotation : action.rotationFrom!
  const delta = action.rotationTo! - from
  const repeatOffset = delta * (action.reversePlayback ? -1 : 1) * completedAnimationRepeats(action, iterationIndex)
  const value = action.rotationKeyframes?.length
    ? animationKeyframeValue(action.rotationKeyframes, action.rotationKeyframeMode, progress)
    : from + delta * progress
  return value + repeatOffset
}

function animationKeyframeValue(keyframes: Array<{ offset: number; value: number }>, mode: 'lin' | 'discrete' | undefined, progress: number): number {
  if (mode === 'discrete') {
    let frame = keyframes[0]!
    for (const candidate of keyframes) {
      if (candidate.offset > progress) break
      frame = candidate
    }
    return frame.value
  }
  const nextIndex = keyframes.findIndex(frame => frame.offset >= progress)
  if (nextIndex === 0) return keyframes[0]!.value
  if (nextIndex < 0) {
    if (keyframes.length < 2) return keyframes.at(-1)!.value
    const from = keyframes.at(-2)!, to = keyframes.at(-1)!
    const segmentProgress = (progress - from.offset) / (to.offset - from.offset)
    return from.value + (to.value - from.value) * segmentProgress
  }
  const from = keyframes[nextIndex - 1]!, to = keyframes[nextIndex]!
  const segmentProgress = (progress - from.offset) / (to.offset - from.offset)
  return from.value + (to.value - from.value) * segmentProgress
}

function animationOpacity(
  element: PptxElement,
  property: NonNullable<PptxAnimation['opacityProperty']> = 'style.opacity',
  paragraphIndex?: number,
  characterStart?: number,
  characterEnd?: number,
): number {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => {
    if (action.effect !== 'opacity' || action.opacityProperty !== property) return false
    const hasTextRange = Boolean(action.paragraphRange || action.characterRange || action.iteration)
    if (paragraphIndex === undefined) return !hasTextRange
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(action, characterStart, characterEnd)
    return hasTextRange
      && (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
      && (!action.characterRange || characterStart !== undefined && characterEnd !== undefined && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  let opacity = 1
  for (const action of actions) {
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(action, characterStart, characterEnd)
    const progress = animationProgress(action, iterationIndex)
    if (progress === -2) continue
    if (progress === null || progress < 0 || action.opacityFrom === undefined || action.opacityTo === undefined) continue
    const formulaSamples = action.opacityFormulaSamples
    const keyframes = action.opacityKeyframes
    let animatedOpacity: number
    if (formulaSamples?.length) animatedOpacity = animationKeyframeValue(formulaSamples, 'lin', progress)
    else if (keyframes?.length) animatedOpacity = animationKeyframeValue(keyframes, action.opacityKeyframeMode, progress)
    else animatedOpacity = action.opacityFrom + (action.opacityTo - action.opacityFrom) * progress
    if (action.opacityAccumulate) {
      // Each completed cycle adds one endpoint delta before composing the animation value.
      const completedRepeats = completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1)
      animatedOpacity += (action.opacityTo - action.opacityFrom) * completedRepeats
    }
    opacity = action.opacityAdditive === 'sum'
      ? opacity + animatedOpacity
      : action.opacityAdditive === 'mult'
        ? opacity * animatedOpacity
        : animatedOpacity
  }
  return Math.max(0, Math.min(1, opacity))
}

function animationOnOff(element: PptxElement, property: NonNullable<PptxAnimation['onOffProperty']>): boolean {
  let enabled = true
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'onOff' && item.onOffProperty === property) || []) {
    const progress = animationProgress(action)
    if (progress === -2) {
      enabled = true
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.onOffKeyframes?.length) {
      for (const frame of action.onOffKeyframes) {
        if (frame.offset > progress) break
        enabled = frame.value
      }
    } else if (action.onOffFrom !== undefined && progress < 1) {
      enabled = action.onOffFrom
    } else if (progress >= 1) {
      enabled = action.onOffTo ?? action.onOffFrom ?? enabled
    }
  }
  return enabled
}

function animationStrokeDash(element: PptxElement): PptxElement['strokeDash'] {
  let strokeDash = element.strokeDash
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'strokeDash') || []) {
    const progress = animationProgress(action)
    if (progress === -2) {
      strokeDash = element.strokeDash
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.strokeDashKeyframes?.length) {
      for (const frame of action.strokeDashKeyframes) {
        if (frame.offset > progress) break
        strokeDash = frame.value
      }
    } else if (action.strokeDashFrom !== undefined && progress < 1) {
      strokeDash = action.strokeDashFrom
    } else if (progress >= 1) {
      strokeDash = action.strokeDashTo ?? action.strokeDashFrom ?? strokeDash
    }
  }
  return strokeDash
}

function animationStrokeWidth(element: PptxElement): number {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'strokeWidth') || []
  let strokeWidth = element.strokeWidth
  for (const action of actions) {
    const progress = animationProgress(action)
    if (progress === -2) {
      strokeWidth = element.strokeWidth
      continue
    }
    if (progress === null || progress < 0 || action.strokeWidthFrom === undefined || action.strokeWidthTo === undefined) continue
    strokeWidth = action.strokeWidthFormulaSamples?.length
      ? animationKeyframeValue(action.strokeWidthFormulaSamples, 'lin', progress)
      : action.strokeWidthKeyframes?.length
        ? animationKeyframeValue(action.strokeWidthKeyframes, action.strokeWidthKeyframeMode, progress)
        : action.strokeWidthFrom + (action.strokeWidthTo - action.strokeWidthFrom) * progress
  }
  // Office zero-width filled outlines are one device pixel, independent of slide zoom.
  // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5564035f-88ff-4ab8-bbfb-d2998a83324d
  if (strokeWidth === 0 && (animationColor(element, 'stroke.color') || element.stroke) !== 'transparent' && scale.value > 0) {
    return 1 / (scale.value * pixelRatio.value)
  }
  return strokeWidth
}

function animationColor(element: PptxElement, property: NonNullable<PptxAnimation['colorProperty']>, paragraphIndex?: number, characterStart?: number, characterEnd?: number, baselineColor?: string): string | undefined {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'color' || action.colorProperty !== property) return false
    const paragraphRange = action.paragraphRange
    const characterRange = action.characterRange
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(action, characterStart, characterEnd)
    return (!paragraphRange || paragraphIndex !== undefined && paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end)
      && (!characterRange || characterStart !== undefined && characterEnd !== undefined && characterStart < characterRange.end && characterEnd > characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  let color: string | undefined
  for (const action of actions) {
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(action, characterStart, characterEnd)
    const progress = animationProgress(action, iterationIndex)
    if (progress === -2) {
      color = undefined
      continue
    }
    let from = action.colorFrom, to = action.colorTo
    if (action.colorFromCurrent) {
      const source = property === 'style.color' ? baselineColor : property === 'fillcolor' ? element.fill : property === 'stroke.color' ? element.stroke : undefined
      const baseline = animationColorChannelsFromCss(source)
      if (baseline) {
        const [hue, saturation, lightness] = rgbToHsl(baseline[0], baseline[1], baseline[2])
        from = action.colorSpace === 'hsl'
          ? [hue, saturation, lightness, baseline[3]]
          : baseline
        if (action.colorBy) {
          const offset = action.colorBy
          to = from.map((channel, index) => {
            if (index === 3) return channel
            const sum = channel + offset[index]!
            return index === 0 && action.colorSpace === 'hsl' ? sum : Math.max(0, Math.min(index === 0 ? 1 : action.colorSpace === 'rgb' && index < 3 ? 255 : 1, sum))
          }) as [number, number, number, number]
        }
      }
    }
    const keyframes = action.colorKeyframes
    if (progress === null || progress < 0 || !from || !to) continue
    let colorProgress: number = progress
    if (keyframes?.length) {
      if (action.colorKeyframeMode === 'discrete') {
        let selected = keyframes[0]!
        for (const frame of keyframes) {
          if (frame.offset > progress) break
          selected = frame
        }
        from = selected.value
        to = selected.value
        colorProgress = 0
      } else {
        let start = keyframes[0]!
        let end = keyframes.at(-1)!
        if (progress <= start.offset) {
          end = start
        } else if (progress < end.offset) {
          for (let index = 1; index < keyframes.length; index++) {
            const next = keyframes[index]!
            if (progress <= next.offset) {
              start = keyframes[index - 1]!
              end = next
              break
            }
          }
        } else {
          start = keyframes.at(-2) || end
        }
        from = start.value
        to = end.value
        colorProgress = end.offset === start.offset ? 0 : (progress - start.offset) / (end.offset - start.offset)
      }
    }
    let channels: [number, number, number, number]
    if (action.colorSpace === 'hsl') {
      const hueDelta = ((to[0] - from[0]) % 1 + 1) % 1
      const huePath = action.colorDirection === 'ccw' && hueDelta ? hueDelta - 1 : hueDelta
      const [r, g, b] = hslToRgb(from[0] + huePath * colorProgress, from[1] + (to[1] - from[1]) * colorProgress, from[2] + (to[2] - from[2]) * colorProgress)
      channels = [r, g, b, from[3] + (to[3] - from[3]) * colorProgress]
    } else {
      channels = [0, 1, 2, 3].map(index => from[index]! + (to[index]! - from[index]!) * colorProgress) as [number, number, number, number]
    }
    color = `rgba(${channels.slice(0, 3).map(channel => Math.round(Math.max(0, Math.min(channel, 255)))).join(', ')}, ${Math.max(0, Math.min(channels[3], 1))})`
  }
  return color
}

function animationColorChannelsFromCss(value?: string): [number, number, number, number] | undefined {
  const color = value?.trim()
  if (!color) return undefined
  const hex = /^#([\da-f]{3}|[\da-f]{4}|[\da-f]{6}|[\da-f]{8})$/i.exec(color)?.[1]
  if (hex) {
    const expand = (channel: string) => channel.length === 1 ? Number.parseInt(channel + channel, 16) : Number.parseInt(channel, 16)
    const channels = hex.length <= 4 ? [...hex].map(expand) : [hex.slice(0, 2), hex.slice(2, 4), hex.slice(4, 6), ...(hex.length === 8 ? [hex.slice(6, 8)] : [])].map(expand)
    return [channels[0]!, channels[1]!, channels[2]!, channels.length === 4 ? channels[3]! / 255 : 1]
  }
  const match = /^rgba?\(\s*(\d+(?:\.\d+)?%?)\s*,\s*(\d+(?:\.\d+)?%?)\s*,\s*(\d+(?:\.\d+)?%?)(?:\s*,\s*(\d+(?:\.\d+)?%?))?\s*\)$/i.exec(color)
  if (!match) return undefined
  const channel = (raw: string) => raw.endsWith('%') ? Number(raw.slice(0, -1)) * 2.55 : Number(raw)
  const alpha = match[4] ? (match[4].endsWith('%') ? Number(match[4].slice(0, -1)) / 100 : Number(match[4])) : 1
  return [channel(match[1]!), channel(match[2]!), channel(match[3]!), alpha]
}

function textColorAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number, baselineColor?: string): CSSProperties {
  const color = animationColor(element, 'style.color', paragraphIndex, characterStart, characterEnd, baselineColor)
  if (!color) return {}
  return { color: animatedRunColor(color, animationOpacity(element, 'fill.opacity', paragraphIndex, characterStart, characterEnd)) }
}

function runHasTextFill(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], paragraphIndex: number, characterStart: number, characterEnd: number): boolean {
  return Boolean(run.fillPattern || run.fillGradient) && !animationColor(element, 'style.color', paragraphIndex, characterStart, characterEnd, run.color)
}

function textRangeOpacityAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties {
  const opacity = animationOpacity(element, 'style.opacity', paragraphIndex, characterStart, characterEnd)
  return opacity === 1 ? {} : { opacity }
}

function textRunAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number, run: PptxElement['paragraphs'][number]['runs'][number], fontScale = element.textFontScale ?? 1, script: ScriptFont = 'latin', separateDecorations = false, filterKey = `${paragraphIndex}-${characterStart}-${characterEnd}`): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'fontSize') return false
    const paragraphRange = action.paragraphRange
    const characterRange = action.characterRange
    const iterationIndex = iterationIndexFor(action, characterStart, characterEnd)
    return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end)
      && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  let factor = 1
  for (const action of actions) {
    const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd))
    if (progress === -2) {
      factor = 1
      continue
    }
    if (progress === null || progress < 0 || action.fontSizeFrom === undefined || action.fontSizeTo === undefined) continue
    factor = action.fontSizeFormulaSamples?.length
      ? animationKeyframeValue(action.fontSizeFormulaSamples, 'lin', progress)
      : action.fontSizeKeyframes?.length
      ? animationKeyframeValue(action.fontSizeKeyframes, action.fontSizeKeyframeMode, progress)
      : action.fontSizeFrom + (action.fontSizeTo - action.fontSizeFrom) * progress
  }
  const baseFontWeight = run.bold ? 700 : 400
  let weight: number | undefined
  const weightActions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'fontWeight') return false
    const paragraphRange = action.paragraphRange
    const characterRange = action.characterRange
    const iterationIndex = iterationIndexFor(action, characterStart, characterEnd)
    return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end)
      && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  for (const action of weightActions) {
    const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd))
    if (progress === -2) {
      weight = baseFontWeight
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.fontWeightReset) {
      weight = baseFontWeight
      continue
    }
    if (action.fontWeightFrom === undefined || action.fontWeightTo === undefined) continue
    const formulaSamples = action.fontWeightFormulaSamplesByBaseline?.find(entry => entry.baseline === baseFontWeight)?.samples || action.fontWeightFormulaSamples
    weight = formulaSamples?.length
      ? animationKeyframeValue(formulaSamples, 'lin', progress)
      : action.fontWeightKeyframes?.length
      ? animationKeyframeValue(action.fontWeightKeyframes, action.fontWeightKeyframeMode, progress)
      : action.fontWeightFrom + (action.fontWeightTo - action.fontWeightFrom) * progress
  }
  let italic: boolean | undefined
  let fontFamily: CSSProperties['fontFamily'] | undefined
  let underline: boolean | undefined
  let lineThrough: boolean | undefined
  let verticalAlign: CSSProperties['verticalAlign'] | undefined
  let verticalAlignChanged = false
  let textTransformScale = 1
  let textShadow: CSSProperties['textShadow']
  let textShadowChanged = false
  let textEmbossEnabled = false
  let textEmbossChanged = false
  const baseFontFamily = script === 'eastAsia' ? run.fontFamilyEastAsia : script === 'complexScript' ? run.fontFamilyComplexScript : run.fontFamily
  let textOutline: string | undefined
  let textOutlineChanged = false
  const baseTextOutline = run.outlineColor !== undefined && run.outlineWidth !== undefined
    ? `${run.outlineWidth * scale.value}px ${run.outlineColor}`
    : '0px transparent'
  const textStyleActions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'textStyle') return false
    const paragraphRange = action.paragraphRange
    const characterRange = action.characterRange
    const iterationIndex = iterationIndexFor(action, characterStart, characterEnd)
    return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end)
      && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  for (const action of textStyleActions) {
    const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd))
    let value: PptxAnimation['textStyleFrom']
    if (progress === -2 && action.textStyleProperty === 'fontWeight') {
      weight = baseFontWeight
      continue
    } else if (progress === -2 && action.textStyleProperty === 'textTransform') {
      verticalAlign = run.baselineOffsetEm === undefined ? 'baseline' : `${run.baselineOffsetEm}em`
      verticalAlignChanged = true
      textTransformScale = 1
      continue
    } else if (progress === -2 && action.textStyleProperty === 'outline') {
      textOutline = baseTextOutline
      textOutlineChanged = true
      continue
    } else if (progress === -2 && action.textStyleProperty === 'textShadow') {
      textShadow = runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd)
      textShadowChanged = true
      continue
    } else if (progress === -2 && action.textStyleProperty === 'emboss') {
      textEmbossEnabled = false
      textEmbossChanged = true
      continue
    } else if (progress === -2) {
      if (action.textStyleProperty === 'fontFamily') value = baseFontFamily || ''
      else if (action.textStyleProperty === 'fontStyle') value = Boolean(run.italic)
      else if (action.textStyleProperty === 'underline') value = Boolean(run.underline)
      else if (action.textStyleProperty === 'lineThrough') value = Boolean(run.lineThrough)
    }
    else if (progress === null || progress < 0) continue
    else if (action.textStyleKeyframes?.length) {
      value = action.textStyleKeyframes[0]!.value
      for (const frame of action.textStyleKeyframes) {
        if (frame.offset > progress) break
        value = frame.value
      }
    } else if (progress >= 1) value = action.textStyleTo
    else value = action.textStyleFrom
    if (value === undefined) continue
    if (action.textStyleProperty === 'fontWeight') weight = value === 'bold' ? 700 : value === 'normal' ? 400 : baseFontWeight
    else if (action.textStyleProperty === 'fontStyle') italic = value as boolean
    else if (action.textStyleProperty === 'fontFamily') fontFamily = cssFontFamily(String(value))
    else if (action.textStyleProperty === 'underline') underline = value as boolean
    else if (action.textStyleProperty === 'lineThrough') lineThrough = value as boolean
    else if (action.textStyleProperty === 'textTransform') {
      verticalAlign = value === 'sub' || value === 'super' ? value : 'baseline'
      verticalAlignChanged = true
      // Match CSS2's conventional sub/superscript size alongside the baseline shift.
      // https://www.w3.org/TR/2008/REC-CSS2-20080411/css2.html#appendix-d
      // https://www.w3.org/TR/CSS2/visudet.html#propdef-vertical-align
      textTransformScale = value === 'sub' || value === 'super' ? 0.83 : 1
    } else if (action.textStyleProperty === 'outline') {
      textOutline = value === true ? baseTextOutline : '0px transparent'
      textOutlineChanged = true
    } else if (action.textStyleProperty === 'textShadow') {
      textShadow = value === 'none' ? 'none' : runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd)
      textShadowChanged = true
    } else if (action.textStyleProperty === 'emboss') {
      textEmbossEnabled = value === 'emboss'
      textEmbossChanged = true
    }
  }
  const embossOffset = Math.max(0.5, (run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value / 12)
  const textDecorationLine = [
    (underline ?? run.underline) ? 'underline' : undefined,
    (lineThrough ?? run.lineThrough) ? 'line-through' : undefined,
  ].filter(Boolean).join(' ')
  const textDecorationStyles: CSSProperties = separateDecorations
    ? {
        ...(underline === undefined ? {} : { '--pptx-underline-line': underline ? 'underline' : 'none' }),
        ...(lineThrough === undefined ? {} : { textDecorationLine: lineThrough ? 'line-through' : 'none' }),
      }
    : underline === undefined && lineThrough === undefined
      ? {}
      : { textDecorationLine: textDecorationLine || 'none' }
  const filter = [
    !textShadowChanged || textShadow !== 'none' ? textInnerShadowFilterReference(element, run, paragraphIndex, characterStart, characterEnd, filterKey) : undefined,
    textEmbossChanged && textEmbossEnabled ? `drop-shadow(-${embossOffset}px -${embossOffset}px 0 rgba(255, 255, 255, 0.65)) drop-shadow(${embossOffset}px ${embossOffset}px 0 rgba(0, 0, 0, 0.65))` : undefined,
  ].filter(Boolean).join(' ')
  const animatedFontSizeFactor = factor * textTransformScale
  return {
    ...(animatedFontSizeFactor === 1 ? {} : { fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value * animatedFontSizeFactor}px` }),
    ...(weight === undefined ? {} : { fontWeight: weight }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(italic === undefined ? {} : { fontStyle: italic ? 'italic' : 'normal' }),
    ...textDecorationStyles,
    ...(verticalAlignChanged ? { verticalAlign } : {}),
    ...(textOutlineChanged ? { WebkitTextStroke: textOutline } : {}),
    ...(textShadowChanged ? { textShadow } : {}),
    ...(textShadowChanged || textEmbossChanged ? { filter: filter || 'none' } : {}),
  }
}

function textTransformAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => {
    const iterationIndex = iterationIndexFor(action, characterStart, characterEnd)
    if (!['scale', 'zoom', 'rotation', 'motion', 'slide'].includes(action.effect) || !(action.paragraphRange || action.characterRange || action.iteration)) return false
    return (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
      && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  if (!actions.length) return {}
  const scaleLayers: AnimationScaleLayers = { dimensions: [1, 1], transform: [1, 1] }
  let rotation = 0
  const [motionX, motionY] = animationMotion(element, paragraphIndex, characterStart, characterEnd)
  const [slideX, slideY] = animationSlide(element, paragraphIndex, characterStart, characterEnd)
  for (const action of actions) {
    const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd))
    if (progress === -2) {
      if (action.effect === 'scale' || action.effect === 'zoom') {
        if (!action.scaleAdditive && !action.scaleMultiplicative) applyAnimationScaleLayer(scaleLayers, action, [1, 1])
      }
      else if (action.effect === 'rotation') rotation = 0
      continue
    }
    if (progress === null || progress < 0) {
      if (action.effect === 'zoom' && action.direction === 'in' && action.scaleFrom) applyAnimationScaleLayer(scaleLayers, action, action.scaleFrom)
      continue
    }
    if ((action.effect === 'scale' || action.effect === 'zoom') && action.scaleFrom && action.scaleTo) {
      const currentScale = action.scaleLayer === 'dimensions' ? scaleLayers.dimensions : scaleLayers.transform
      const animatedScale = animationScaleValue(action, progress, iterationIndexFor(action, characterStart, characterEnd), action.scaleFromCurrent ? currentScale : undefined)
      applyAnimationScaleLayer(scaleLayers, action, animatedScale)
    } else if (action.effect === 'rotation' && action.rotationFrom !== undefined && action.rotationTo !== undefined) {
      const angle = animationRotationValue(action, progress, rotation, iterationIndexFor(action, characterStart, characterEnd))
      rotation = action.rotationRelative || action.rotationAdditive ? rotation + angle : angle
    }
  }
  const scaling = composedAnimationScale(scaleLayers)
  const transforms = [
    motionX || motionY ? `translate(${motionX * scale.value}px, ${motionY * scale.value}px)` : '',
    slideX || slideY ? `translate(${slideX * scale.value}px, ${slideY * scale.value}px)` : '',
    scaling[0] !== 1 || scaling[1] !== 1 ? `scale(${scaling[0]}, ${scaling[1]})` : '',
    rotation ? `rotate(${rotation}deg)` : '',
  ].filter(Boolean)
  return { display: 'inline-block', transform: transforms.join(' '), transformOrigin: 'center center' }
}

function animationMotion(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): [number, number] {
  const textRange = paragraphIndex !== undefined && characterStart !== undefined && characterEnd !== undefined
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'motion'
    && (textRange
      ? Boolean(action.paragraphRange || action.characterRange || action.iteration)
        && (!action.paragraphRange || paragraphIndex! >= action.paragraphRange.start && paragraphIndex! <= action.paragraphRange.end)
        && (!action.characterRange || characterStart! < action.characterRange.end && characterEnd! > action.characterRange.start)
      : !action.paragraphRange && !action.characterRange && !action.iteration)) || []
  let value: [number, number] = [0, 0]
  const applyMotionValue = (action: RenderAnimation, motion: [number, number]) => {
    if (action.motionCoordinateProperties) {
      const mapped: [number, number] = [0, 0]
      for (const [index, property] of action.motionCoordinateProperties.entries()) {
        const axis = property === 'ppt_x' ? 0 : 1
        mapped[axis] += motion[index] || 0
      }
      value = action.motionAdditive ? [value[0] + mapped[0], value[1] + mapped[1]] : mapped
    } else if (action.motionCoordinateProperty === 'ppt_x') value = [action.motionAdditive ? value[0] + motion[0] : motion[0], value[1]]
    else if (action.motionCoordinateProperty === 'ppt_y') value = [value[0], action.motionAdditive ? value[1] + motion[1] : motion[1]]
    else value = action.motionAdditive
      ? [value[0] + motion[0], value[1] + motion[1]]
      : motion
  }
  for (const action of actions) {
    const iterationIndex = textRange ? iterationIndexFor(action, characterStart!, characterEnd!) : undefined
    const progress = animationProgress(action, iterationIndex)
    if (progress === -2) {
      if (!action.motionAdditive) {
        if (action.motionCoordinateProperty === 'ppt_x') value = [0, value[1]]
        else if (action.motionCoordinateProperty === 'ppt_y') value = [value[0], 0]
        else value = [0, 0]
      }
      continue
    }
    const samples = action.motionPathSamples
    if (progress !== null && progress >= 0 && samples?.length) {
      if (samples.length === 1) {
        const offset = motionAccumulationOffset(action, iterationIndex)
        applyMotionValue(action, [samples[0]!.x + offset[0], samples[0]!.y + offset[1]])
        continue
      }
      const samplePosition = progress * (samples.length - 1)
      const fromIndex = Math.max(0, Math.min(samples.length - 2, Math.floor(samplePosition)))
      const toIndex = Math.min(samples.length - 1, fromIndex + 1)
      const from = samples[fromIndex]!
      const to = samples[toIndex]!
      const ratio = samplePosition - fromIndex
      const offset = motionAccumulationOffset(action, iterationIndex)
      const motion: [number, number] = [from.x + (to.x - from.x) * ratio + offset[0], from.y + (to.y - from.y) * ratio + offset[1]]
      applyMotionValue(action, motion)
      continue
    }
    const path = action.motionPath
    if (progress === null || progress < 0 || !path?.length) continue
    const offset = motionAccumulationOffset(action, iterationIndex)
    if (path.length === 1) {
      const point = path[0]
      if (point) applyMotionValue(action, [point.x + offset[0], point.y + offset[1]])
      continue
    }
    const distance = (action.motionPathLength || 0) * progress
    if (distance > (action.motionPathLength || 0) && path.length > 1) {
      const from = path.at(-2)!, to = path.at(-1)!
      const span = to.distance - from.distance
      const ratio = span ? (distance - from.distance) / span : 1
      const offset = motionAccumulationOffset(action, iterationIndex)
      applyMotionValue(action, [from.x + (to.x - from.x) * ratio + offset[0], from.y + (to.y - from.y) * ratio + offset[1]])
      continue
    }
    let low = 0, high = path.length
    while (low < high) {
      const middle = (low + high) >>> 1
      const point = path[middle]
      if (point && point.distance <= distance) low = middle + 1
      else high = middle
    }
    const from = path[low - 1], to = path[low]
    if (!from) continue
    if (!to) {
      applyMotionValue(action, [from.x + offset[0], from.y + offset[1]])
      continue
    }
    const span = to.distance - from.distance
    const ratio = span ? (distance - from.distance) / span : 0
    applyMotionValue(action, [from.x + (to.x - from.x) * ratio + offset[0], from.y + (to.y - from.y) * ratio + offset[1]])
  }
  return value
}

function animationSlide(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): [number, number] {
  let value: [number, number] = [0, 0]
  const textRange = paragraphIndex !== undefined || characterStart !== undefined
  const actions = animationsByTarget.value.get(element.id)?.filter(item => {
    if (item.effect !== 'slide') return false
    const iterationIndex = characterStart !== undefined && characterEnd !== undefined
      ? iterationIndexFor(item, characterStart, characterEnd)
      : undefined
    if (!textRange) return !item.paragraphRange && !item.characterRange && !item.iteration
    return (Boolean(item.paragraphRange || item.characterRange) || iterationIndex !== undefined && iterationIndex >= 0)
      && (!item.paragraphRange || paragraphIndex !== undefined && paragraphIndex >= item.paragraphRange.start && paragraphIndex <= item.paragraphRange.end)
      && (!item.characterRange || characterStart !== undefined && characterEnd !== undefined && characterStart < item.characterRange.end && characterEnd > item.characterRange.start)
      && (!item.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  }) || []
  for (const action of actions) {
    const iterationIndex = characterStart !== undefined && characterEnd !== undefined
      ? iterationIndexFor(action, characterStart, characterEnd)
      : undefined
    const progress = animationProgress(action, iterationIndex)
    if (progress === -2) { value = [0, 0]; continue }
    if (progress === null) continue
    const phase = progress < 0 ? 0 : progress
    const distance = action.direction === 'in' ? 1 - phase : phase
    value = action.slideFrom === 'left' ? [-props.width * distance, 0]
      : action.slideFrom === 'right' ? [props.width * distance, 0]
        : action.slideFrom === 'top' ? [0, -props.height * distance]
          : action.slideFrom === 'bottom' ? [0, props.height * distance]
            : action.slideFrom === 'topLeft' ? [-props.width * distance, -props.height * distance]
              : action.slideFrom === 'topRight' ? [props.width * distance, -props.height * distance]
                : action.slideFrom === 'bottomLeft' ? [-props.width * distance, props.height * distance]
                  : action.slideFrom === 'bottomRight' ? [props.width * distance, props.height * distance]
            : [0, 0]
  }
  return value
}

function animationTargetsRange(action: PptxAnimation, paragraphIndex?: number, characterStart?: number, characterEnd?: number): boolean {
  // p:iterate ranges are text offsets and can only be resolved for character-level calls.
  if (characterStart !== undefined && characterEnd !== undefined) {
    const overlaps = Boolean(action.characterRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
    const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : undefined
    return overlaps || (iterationIndex !== undefined && iterationIndex >= 0)
  }
  if (paragraphIndex !== undefined) {
    return Boolean(action.paragraphRange && !action.characterRange && !action.iteration && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
  }
  return !action.paragraphRange && !action.characterRange && !action.iteration
}

function animationIterationIndex(action: PptxAnimation, paragraphIndex?: number, characterStart?: number, characterEnd?: number): number | undefined {
  if (characterStart !== undefined && characterEnd !== undefined) return iterationIndexFor(action, characterStart, characterEnd)
  return paragraphIndex === undefined ? undefined : iterationIndexFor(action, paragraphIndex, paragraphIndex + 1)
}

function animationClipMaskStyle(element: PptxElement, clipPath: string, textRange: boolean): CSSProperties {
  if (!textRange) return { clipPath }
  const { width, height } = elementRenderSize(element)
  const mask = clipPathMaskImage(clipPath, width * scale.value, height * scale.value)
  if (!mask) return { clipPath }
  return {
    maskImage: mask, WebkitMaskImage: mask,
    maskSize: '100% 100%', WebkitMaskSize: '100% 100%',
    maskRepeat: 'no-repeat', WebkitMaskRepeat: 'no-repeat',
  }
}

function animationWipe(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'wipe') return false
    if (characterStart !== undefined && characterEnd !== undefined) {
      const overlaps = Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
      const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : undefined
      return overlaps || (iterationIndex !== undefined && iterationIndex >= 0)
    }
    return paragraphIndex === undefined
      ? !action.paragraphRange && !action.characterRange && !action.iteration
      : Boolean(action.paragraphRange && !action.iteration && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
  }) || []
  const first = actions[0]
  if (!first?.wipeDirection) return {}
  let progress = first.direction === 'in' ? 0 : 100
  let direction = first.wipeDirection
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd))
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    direction = action.wipeDirection || direction
    progress = action.direction === 'in' ? phase * 100 : (1 - phase) * 100
  }
  if (reset) return {}
  const cssDirection = direction === 'up' ? 'bottom' : direction === 'down' ? 'top' : direction
  const mask = `linear-gradient(to ${cssDirection}, #000 var(--wipe-progress), transparent var(--wipe-progress))`
  return { '--wipe-progress': `${progress}%`, maskImage: mask, WebkitMaskImage: mask } as CSSProperties
}

function animationBlinds(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => {
    if (action.effect !== 'blinds') return false
    if (paragraphIndex !== undefined) {
      return Boolean(!action.iteration && action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
    }
    if (characterStart !== undefined && characterEnd !== undefined) {
      const overlaps = Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
      const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : undefined
      const iteratesCharacter = iterationIndex !== undefined && iterationIndex >= 0
      return overlaps || iteratesCharacter
    }
    return !action.characterRange && !action.paragraphRange && !action.iteration
  }) || []
  const first = actions[0]
  if (!first?.blindsOrientation) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let orientation = first.blindsOrientation
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd))
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    orientation = action.blindsOrientation || orientation
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  const stripeSize = `${Math.max(0, Math.min(progress, 1) * 12.5)}%`
  const maskDirection = orientation === 'horizontal' ? 'to bottom' : 'to right'
  const mask = `repeating-linear-gradient(${maskDirection}, #000 0 ${stripeSize}, transparent ${stripeSize} 12.5%)`
  return { '--blinds-stripe-size': stripeSize, maskImage: mask, WebkitMaskImage: mask } as CSSProperties
}

function animationChecker(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'checker') return false
    if (paragraphIndex !== undefined) return Boolean(action.paragraphRange && !action.characterRange && !action.iteration && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
    if (characterStart !== undefined && characterEnd !== undefined) {
      const overlaps = Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
      const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : undefined
      return overlaps || (iterationIndex !== undefined && iterationIndex >= 0)
    }
    return !action.paragraphRange && !action.characterRange && !action.iteration
  }) || []
  const first = actions[0]
  if (!first?.checkerOrientation) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let orientation = first.checkerOrientation
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd))
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    orientation = action.checkerOrientation || orientation
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  const { width: localWidth, height: localHeight } = elementRenderSize(element)
  const width = localWidth * scale.value
  const height = localHeight * scale.value
  const amount = Math.max(0, Math.min(progress, 1))
  if (paragraphIndex === undefined && characterStart === undefined) return { clipPath: gridTransitionClipPath(width, height, orientation, amount, true) }
  const mask = gridTransitionMaskImage(width, height, orientation, amount, true)
  return {
    maskImage: mask, WebkitMaskImage: mask,
    maskSize: '100% 100%', WebkitMaskSize: '100% 100%',
    maskRepeat: 'no-repeat', WebkitMaskRepeat: 'no-repeat',
  }
}

function animationShapeMask(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'shape'
    && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || []
  const first = actions[0]
  if (!first?.shapeFilter || !first.shapeDirection) return {}
  let current: { shape: 'circle' | 'diamond' | 'box' | 'plus'; direction: 'in' | 'out'; progress: number } = {
    shape: first.shapeFilter,
    direction: first.shapeDirection,
    progress: first.direction === 'in' ? 0 : 1,
  }
  let reset = false
  for (const action of actions) {
    if (!action.shapeFilter || !action.shapeDirection) continue
    const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd))
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    current = {
      shape: action.shapeFilter,
      direction: action.shapeDirection,
      progress: action.direction === 'in' ? phase : 1 - phase,
    }
  }
  if (reset || !current) return {}
  const { width, height } = elementRenderSize(element)
  const clipPath = shapeEffectClipPath(current.shape, width * scale.value, height * scale.value, current.progress, current.direction)
  return animationClipMaskStyle(element, clipPath, paragraphIndex !== undefined || characterStart !== undefined)
}

function animationDissolve(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'dissolve'
    && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || []
  const first = actions[0]
  if (!first) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd))
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  const { width, height } = elementRenderSize(element)
  const clipPath = dissolveClipPath(width * scale.value, height * scale.value, progress, element.id)
  return animationClipMaskStyle(element, clipPath, paragraphIndex !== undefined || characterStart !== undefined)
}

function animationWheel(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'wheel'
    && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || []
  const first = actions[0]
  if (!first?.wheelSpokes) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let spokes = first.wheelSpokes
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd))
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    spokes = action.wheelSpokes || spokes
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  const { width, height } = elementRenderSize(element)
  const clipPath = wheelClipPath(width * scale.value, height * scale.value, spokes, progress)
  return animationClipMaskStyle(element, clipPath, paragraphIndex !== undefined || characterStart !== undefined)
}

function animationRandomBars(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'randomBars'
    && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || []
  const first = actions[0]
  if (!first?.randomBarOrientation) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let orientation = first.randomBarOrientation
  let reset = false
  for (const action of actions) {
    const iterationIndex = animationIterationIndex(action, paragraphIndex, characterStart, characterEnd)
    const phase = animationProgress(action, iterationIndex)
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    orientation = action.randomBarOrientation || orientation
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  const { width, height } = elementRenderSize(element)
  const clipPath = randomBarClipPath(width * scale.value, height * scale.value, orientation, progress, element.id)
  return animationClipMaskStyle(element, clipPath, paragraphIndex !== undefined || characterStart !== undefined)
}

function animationStrips(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'strips'
    && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || []
  const first = actions[0]
  if (!first?.stripsDirection) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let direction = first.stripsDirection
  let reset = false
  for (const action of actions) {
    const iterationIndex = animationIterationIndex(action, paragraphIndex, characterStart, characterEnd)
    const phase = animationProgress(action, iterationIndex)
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    direction = action.stripsDirection || direction
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  const { width, height } = elementRenderSize(element)
  const clipPath = stripsClipPath(width * scale.value, height * scale.value, direction, progress)
  return animationClipMaskStyle(element, clipPath, paragraphIndex !== undefined || characterStart !== undefined)
}

function animationBarn(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'barn'
    && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || []
  const first = actions[0]
  if (!first?.barnOrientation || !first.barnMotion) return {}
  let orientation = first.barnOrientation
  let motion = first.barnMotion
  let progress = 0
  let reset = false
  for (const action of actions) {
    if (!action.barnOrientation || !action.barnMotion) continue
    const iterationIndex = animationIterationIndex(action, paragraphIndex, characterStart, characterEnd)
    const phase = animationProgress(action, iterationIndex)
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    orientation = action.barnOrientation
    motion = action.barnMotion
    progress = phase
  }
  if (reset) return {}
  return { clipPath: barnClipPath(orientation, motion, progress) }
}

function animationCss(actions: RenderAnimation[] | undefined, characterStart?: number, characterEnd?: number): CSSProperties {
  if (props.thumbnail && !props.playbackPreview || !actions?.length) return {}
  const visibilityActions = actions.filter(action => ['appear', 'fade', 'wipe', 'strips', 'barn', 'slide', 'visibility', 'zoom'].includes(action.effect) || action.randomEffectFallback)
  const first = visibilityActions[0]
  if (!first) return {}
  let opacity = first.effect === 'visibility' || first.randomEffectFallback || first.effect === 'wipe' || first.effect === 'strips' || first.effect === 'barn' || first.effect === 'slide' || first.direction === 'out' ? 1 : 0
  let visibility: 'visible' | 'hidden' = first.effect === 'visibility' ? first.visibilityFrom || 'visible' : first.direction === 'out' ? 'visible' : 'hidden'
  for (const action of visibilityActions) {
    const iterationIndex = characterStart === undefined || characterEnd === undefined ? undefined : iterationIndexFor(action, characterStart, characterEnd)
    const progress = animationProgress(action, iterationIndex)
    if (progress === -2) {
      opacity = 1
      visibility = 'visible'
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.effect === 'appear') {
      const atEnd = progress >= 1
      visibility = action.direction === 'in'
        ? atEnd ? 'visible' : 'hidden'
        : atEnd ? 'hidden' : 'visible'
      opacity = visibility === 'visible' ? 1 : 0
    } else if (action.effect === 'visibility') {
      let value = action.visibilityFrom || 'visible'
      if (action.visibilityKeyframes?.length) {
        for (const frame of action.visibilityKeyframes) {
          if (frame.offset > progress) break
          value = frame.value
        }
      } else if (progress >= 1) value = action.visibilityTo || value
      visibility = value
    } else if (action.randomEffectFallback) {
      visibility = action.direction === 'in'
        ? progress > 0 ? 'visible' : 'hidden'
        : progress >= 1 ? 'hidden' : 'visible'
      opacity = action.effect === 'fade'
        ? action.direction === 'in' ? progress : 1 - progress
        : 1
    } else if (action.effect === 'wipe' || action.effect === 'strips' || action.effect === 'barn') {
      visibility = action.direction === 'in'
        ? progress > 0 ? 'visible' : 'hidden'
        : progress >= 1 ? 'hidden' : 'visible'
      opacity = 1
    } else if (action.effect === 'slide' || action.effect === 'zoom') {
      visibility = action.direction === 'in' || progress < 1 ? 'visible' : 'hidden'
      opacity = 1
    } else {
      // p:progress defaults to a linear ramp; p:cTn timing has already shaped progress.
      opacity = action.direction === 'in' ? progress : 1 - progress
      visibility = action.direction === 'in' || progress < 1 ? 'visible' : 'hidden'
    }
  }
  return {
    opacity,
    visibility,
    pointerEvents: visibility === 'visible' && opacity > 0 ? 'auto' : 'none',
  }
}

function animationStyle(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)
  if (!actions?.length) return {}
  const animatedStyle = animationCss(actions.filter(action => !action.paragraphRange && !action.characterRange && !action.iteration))
  const opacity = animationOpacity(element)
  return {
    ...animatedStyle,
    opacity: (typeof animatedStyle.opacity === 'number' ? animatedStyle.opacity : 1) * opacity,
    pointerEvents: opacity === 0 ? 'none' : animatedStyle.pointerEvents,
    ...animationWipe(element),
    ...animationBlinds(element),
    ...animationChecker(element),
    ...animationShapeMask(element),
    ...animationDissolve(element),
    ...animationWheel(element),
    ...animationRandomBars(element),
    ...animationStrips(element),
    ...animationBarn(element),
  }
}

function paragraphAnimationStyle(element: PptxElement, index: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    const range = action.paragraphRange
    return range && action.effect !== 'barn' && !action.iteration && index >= range.start && index <= range.end
  })
  const [slideX, slideY] = animationSlide(element, index)
  return {
    ...animationCss(actions),
    ...(slideX || slideY ? { transform: `translate(${slideX * scale.value}px, ${slideY * scale.value}px)` } : {}),
    ...animationWipe(element, index),
    ...animationBlinds(element, index),
    ...animationChecker(element, index),
    ...animationBarn(element, index),
    ...animationShapeMask(element, index),
    ...animationDissolve(element, index),
    ...animationWheel(element, index),
    ...animationRandomBars(element, index),
    ...animationStrips(element, index),
  }
}

function characterAnimationStyle(element: PptxElement, paragraphIndex: number, start: number, end: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    const range = action.characterRange
    const overlaps = range && start < range.end && end > range.start
    if (action.iteration) return Boolean(overlaps || !range) && (iterationIndexFor(action, start, end) ?? -1) >= 0
    return Boolean(overlaps)
  })
  const animatedStyle = animationCss(actions?.filter(action => action.effect !== 'barn'), start, end)
  const textOpacity = animationOpacity(element, 'style.opacity', paragraphIndex, start, end)
  const opacity = typeof animatedStyle.opacity === 'number' ? animatedStyle.opacity * textOpacity : undefined
  return {
    ...animatedStyle,
    ...(opacity === undefined ? {} : { opacity }),
    ...animationWipe(element, undefined, start, end),
    ...animationBlinds(element, undefined, start, end),
    ...animationChecker(element, undefined, start, end),
    ...animationBarn(element, undefined, start, end),
    ...animationShapeMask(element, undefined, start, end),
    ...animationDissolve(element, undefined, start, end),
    ...animationWheel(element, undefined, start, end),
    ...animationRandomBars(element, undefined, start, end),
    ...animationStrips(element, undefined, start, end),
  }
}

function hasTrigger(elementId: string) {
  return triggerTargetIds.value.has(elementId)
}

function triggerSequenceIsActive(index: number, sequence: NonNullable<PptxSlide['triggeredAnimations']>[number], now: number) {
  return sequence.steps.some((actions, stepIndex) => {
    const key = triggerKey(index, stepIndex)
    const startedAt = triggerStartedAt.get(key)
    const endedAt = triggerStepEndedAt.get(key)
    if (startedAt === undefined || !Number.isFinite(startedAt)) return false
    if (endedAt !== undefined && now >= endedAt) return false
    return actions.some(action => {
      if (!playbackActionEnabled(action)) return false
      const actionStart = startedAt + action.delayMs
      if (now < actionStart) return true
      if (action.effect === 'media' && action.mediaActionKey) {
        if (completedMediaActions.has(action.mediaActionKey)) return false
        if (action.mediaWaitForEnd) {
          if (blockedMediaPlaybackIds.value.includes(action.targetId)) return false
          const media = Array.from(stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || [])
            .find(item => item.dataset.mediaId === action.targetId)
          return !media || !media.paused
        }
      }
      return now < actionStart + naturalAnimationDuration(action)
    })
  })
}

function startTriggeredAnimations(indices: number[], startedAt = performance.now(), refresh = true) {
  if (props.triggerPlayback !== undefined) return
  const sequences = props.slide.triggeredAnimations || []
  if (!indices.length) return
  for (const index of indices) {
    const sequence = sequences[index]
    if (!sequence || !triggerSequenceHasEnabledAction(sequence)) continue
    const count = sequence.steps.length
    if (!count) continue
    let step = triggerSteps.get(index) || 0
    if (step >= count) {
      if (sequence.restart === 'never') continue
      if (sequence.restart === 'whenNotActive' && triggerSequenceIsActive(index, sequence, performance.now())) continue
      for (let previous = 0; previous < count; previous++) {
        const key = triggerKey(index, previous)
        triggerStartedAt.delete(key)
        triggerStepEndedAt.delete(key)
        triggerPlaybackByKey.delete(key)
      }
      timelineActions.value.forEach((action, actionIndex) => {
        if (action.triggerSequence === index) executedMediaActions.delete(actionIndex)
      })
      step = 0
    }
    if (step > 0 && !sequence.concurrent) endTimelineStep(index, step - 1, startedAt)
    const key = triggerKey(index, step)
    const playback = { sequence: index, step, startedAt: performance.timeOrigin + startedAt }
    triggerStartedAt.set(key, startedAt)
    triggerPlaybackByKey.set(key, playback)
    triggerSteps.set(index, step + 1)
  }
  if (refresh) refreshTimeline()
  emit('trigger-playback', [...triggerPlaybackByKey.values()].sort((a, b) => a.sequence - b.sequence || a.step - b.step))
}

function runTimingNodeTriggeredAnimations(now: number): number {
  if (props.thumbnail || props.triggerPlayback !== undefined) return Infinity
  let next = Infinity
  props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
    if (!triggerSequenceHasEnabledAction(sequence)) return
    sequence.conditions.forEach((condition, conditionIndex) => {
      const key = triggerConditionKey(sequenceIndex, conditionIndex)
      if (condition.type === 'timeNode' && !firedTimingNodeTriggers.has(key)) {
        const mediaActionTrigger = condition.sourceMediaActionKey !== undefined
        const sourceStartedAt = mediaActionTrigger ? undefined : condition.sourceStep === undefined ? automaticStartedAt : stepStartedAt.get(condition.sourceStep)
        const mediaEventAt = mediaActionTrigger
          ? condition.event === 'begin'
            ? scheduledMediaActions.get(condition.sourceMediaActionKey!)
            : condition.event === 'onBegin'
              ? startedMediaActions.get(condition.sourceMediaActionKey!)
              : condition.event === 'end'
                ? completedMediaActions.get(condition.sourceMediaActionKey!)
                : endedMediaTargets.get(condition.sourceMediaActionKey!)
          : undefined
        const sourceEndedAt = !mediaActionTrigger && condition.event === 'onEnd' && condition.sourceStep !== undefined ? mainStepEndedAt.get(condition.sourceStep) : undefined
        const sourceTime = mediaActionTrigger
          ? mediaEventAt
          : sourceStartedAt !== undefined && sourceStartedAt !== Number.NEGATIVE_INFINITY
            ? sourceEndedAt ?? sourceStartedAt + condition.sourceDelayMs
            : undefined
        if (sourceTime !== undefined) {
          const startedAt = sourceTime + condition.delayMs
          if (Number.isFinite(startedAt)) {
            if (startedAt > now) next = Math.min(next, startedAt)
            else {
              firedTimingNodeTriggers.add(key)
              triggerConditionDueAt.set(key, startedAt)
              pendingTriggerConditions.add(key)
            }
          }
        }
      }
      const dueAt = triggerConditionDueAt.get(key)
      if (pendingTriggerConditions.has(key) && dueAt !== undefined && dueAt > now) next = Math.min(next, dueAt)
    })

    const pending = sequence.conditions.flatMap((_, conditionIndex) => {
      const key = triggerConditionKey(sequenceIndex, conditionIndex)
      const dueAt = triggerConditionDueAt.get(key)
      return pendingTriggerConditions.has(key) && dueAt !== undefined && dueAt <= now ? [{conditionIndex, dueAt}] : []
    })
    if (!pending.length) return
    const progress = triggerSteps.get(sequenceIndex) || 0
    const advancing = progress > 0 && progress < sequence.steps.length
    const conditionTimes = sequence.conditions.map((_, conditionIndex) => triggerConditionDueAt.get(triggerConditionKey(sequenceIndex, conditionIndex)))
    const allConditionsDue = conditionTimes.every(dueAt => dueAt !== undefined && dueAt <= now)
    if (!advancing && !allConditionsDue) return

    const startedAt = advancing
      ? Math.min(...pending.map(condition => condition.dueAt))
      : Math.max(...conditionTimes as number[])
    sequence.conditions.forEach((condition, conditionIndex) => {
      const key = triggerConditionKey(sequenceIndex, conditionIndex)
      pendingTriggerConditions.delete(key)
      if (condition.type !== 'timeNode') triggerConditionDueAt.delete(key)
    })
    startTriggeredAnimations([sequenceIndex], startedAt, false)
  })
  return next
}

function satisfyTriggeredAnimationConditions(matches: (condition: PptxAnimationTriggerCondition, sequenceIndex: number, conditionIndex: number) => boolean) {
  if (props.thumbnail || props.triggerPlayback !== undefined) return
  const now = performance.now()
  let matched = false
  props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
    if (!triggerSequenceHasEnabledAction(sequence)) return
    sequence.conditions.forEach((condition, conditionIndex) => {
      if (!matches(condition, sequenceIndex, conditionIndex)) return
      matched = true
      const key = triggerConditionKey(sequenceIndex, conditionIndex)
      if (triggerConditionDueAt.has(key)) return
      triggerConditionDueAt.set(key, now + condition.delayMs)
      pendingTriggerConditions.add(key)
    })
  })
  if (matched) {
    animationClock.value = now
    scheduleTimelineUpdate()
  }
}

function runTriggeredAnimations(event: Event, elementId: string, triggerEvent: 'onClick' | 'onDblClick' | 'onMouseOver' | 'onMouseOut') {
  if (props.thumbnail || props.triggerPlayback !== undefined) return
  const sequences = props.slide.triggeredAnimations || []
  const matching = sequences.some(sequence => triggerSequenceHasEnabledAction(sequence)
    && sequence.conditions.some(condition => condition.type === 'shape' && condition.id === elementId && condition.event === triggerEvent))
  if (!matching) {
    if (sequences.some(sequence => triggerSequenceHasEnabledAction(sequence) && sequence.conditions.some(condition => condition.type === 'shape' && condition.id === elementId))) event.stopPropagation()
    return
  }
  if (triggerEvent === 'onClick' || triggerEvent === 'onDblClick') event.stopPropagation()
  satisfyTriggeredAnimationConditions(condition => condition.type === 'shape' && condition.id === elementId && condition.event === triggerEvent)
}

function runSlideTriggeredAnimations(triggerEvent: 'onNext' | 'onPrev' | 'onStopAudio') {
  if (props.thumbnail || props.triggerPlayback !== undefined) return
  satisfyTriggeredAnimationConditions(condition => condition.type === 'slide' && condition.event === triggerEvent)
}

function runMediaBookmarkAnimations(media: HTMLMediaElement, element: PptxElement, mayTrigger: boolean) {
  const currentTime = media.currentTime
  const previousTime = lastMediaBookmarkTimes.get(media)
  if (media.seeking && previousTime !== undefined && currentTime < previousTime) mediaBookmarkPlaybackStarted.delete(media)
  const startingAtBeginning = mayTrigger
    && !media.paused
    && !media.seeking
    && !mediaBookmarkPlaybackStarted.has(media)
    && (previousTime === undefined || previousTime === 0)
    && (element.mediaTrim?.startMs || 0) === 0
  lastMediaBookmarkTimes.set(media, currentTime)
  if (mayTrigger && !media.paused && !media.seeking) mediaBookmarkPlaybackStarted.add(media)
  if (!mayTrigger) return
  const matching = new Set((mediaBookmarkTriggersByMedia.value.get(element.id) || [])
    .filter(trigger => {
      const bookmarkTime = trigger.timeMs / 1000
      return (previousTime ?? 0) < bookmarkTime && currentTime >= bookmarkTime
        || startingAtBeginning && bookmarkTime === 0
    })
    .map(trigger => triggerConditionKey(trigger.sequenceIndex, trigger.conditionIndex)))
  if (matching.size) satisfyTriggeredAnimationConditions((_, sequenceIndex, conditionIndex) => matching.has(triggerConditionKey(sequenceIndex, conditionIndex)))
}

function canActivateHyperlink(hyperlink?: PptxHyperlink): hyperlink is PptxHyperlink {
  return Boolean(hyperlink) && !props.thumbnail && !props.playbackPreview && props.triggerPlayback === undefined
}

function hyperlinkHref(hyperlink?: PptxHyperlink): string {
  return hyperlink?.kind === 'external' ? hyperlink.url : '#slide-link'
}

function onHyperlinkClick(event: Event, hyperlink?: PptxHyperlink) {
  if (!canActivateHyperlink(hyperlink)) return
  event.preventDefault()
  event.stopPropagation()
  emit('activate-hyperlink', hyperlink)
}

function onElementClick(event: Event, elementId: string, hyperlink?: PptxHyperlink) {
  if (canActivateHyperlink(hyperlink)) return
  runTriggeredAnimations(event, elementId, 'onClick')
}

function onElementDoubleClick(event: Event, elementId: string) {
  runTriggeredAnimations(event, elementId, 'onDblClick')
}

function onElementMouseEnter(event: Event, elementId: string) {
  runTriggeredAnimations(event, elementId, 'onMouseOver')
}

function onElementMouseLeave(event: Event, elementId: string) {
  runTriggeredAnimations(event, elementId, 'onMouseOut')
}

function onElementKeyActivate(event: Event, elementId: string, hyperlink?: PptxHyperlink) {
  if (canActivateHyperlink(hyperlink)) {
    if (event instanceof KeyboardEvent && event.key === 'Enter') onHyperlinkClick(event, hyperlink)
    return
  }
  const condition = props.slide.triggeredAnimations?.flatMap(sequence => sequence.conditions).find(item => item.type === 'shape' && item.id === elementId)
  runTriggeredAnimations(event, elementId, condition?.type === 'shape' ? condition.event : 'onClick')
}

function usesRightToLeftColumns(element: PptxElement): boolean {
  return element.textColumnsRightToLeft === true
    && (element.textColumnCount ?? 1) > 1
    && (!element.textOrientation || element.textOrientation === 'horz')
}

function hasRenderableText(element: PptxElement): boolean {
  return element.paragraphs.some(paragraph => paragraph.runs.some(run => run.text.length > 0)
    || paragraph.bullet?.kind === 'char' && paragraph.bullet.char.length > 0
    || paragraph.bullet?.kind === 'number'
    || paragraph.bullet?.kind === 'image' && Boolean(paragraph.bullet.url))
}

function textFrameStyle(element: PptxElement): CSSProperties {
  const { scaleX, scaleY } = elementRenderSize(element)
  const verticalText = element.textOrientation === 'vert' || element.textOrientation === 'vert270'
  const eastAsianVerticalText = element.textOrientation === 'eaVert'
  const wordArtVerticalText = element.textOrientation === 'wordArtVert'
  const wordArtVerticalRtlText = element.textOrientation === 'wordArtVertRtl'
  const mongolianVerticalText = element.textOrientation === 'mongolianVert'
  const columnCount = element.textColumnCount ?? 1
  const verticalColumnWritingMode = columnCount > 1 && element.textOrientation && element.textOrientation !== 'horz' && !wordArtVerticalRtlText
    ? element.textColumnsRightToLeft ? 'vertical-rl' : 'vertical-lr'
    : undefined
  const rect = element.customTextRect
  const fittedSize = rect ? autoFitDimensions.get(`${props.slide.id}:${element.id}`) : undefined
  const ellipsisLineCount = supportsTextVerticalEllipsis(element) ? textEllipsisLineCounts.get(`${props.slide.id}:${element.id}`) : undefined
  const margins = rect ? { ...element.margins } : element.margins
  if (rect) {
    const availableWidth = Math.max(0, (rect.right - rect.left) * (fittedSize?.width || element.width))
    const horizontalMargins = margins.left + margins.right
    if (horizontalMargins > availableWidth) {
      const ratio = availableWidth / horizontalMargins
      margins.left *= ratio
      margins.right *= ratio
    }
    const availableHeight = Math.max(0, (rect.bottom - rect.top) * (fittedSize?.height || element.height))
    const verticalMargins = margins.top + margins.bottom
    if (verticalMargins > availableHeight) {
      const ratio = availableHeight / verticalMargins
      margins.top *= ratio
      margins.bottom *= ratio
    }
  }
  return {
    position: rect ? 'absolute' : undefined,
    left: rect ? `${rect.left / scaleX * 100}%` : undefined,
    top: rect ? `${rect.top / scaleY * 100}%` : undefined,
    width: `${(rect ? (rect.right - rect.left) * 100 : 100) / scaleX}%`,
    height: `${(rect ? (rect.bottom - rect.top) * 100 : 100) / scaleY}%`,
    transform: element.renderMatrix ? `scale(${scaleX},${scaleY})` : undefined,
    transformOrigin: element.renderMatrix ? '0 0' : undefined,
    boxSizing: rect ? 'border-box' : undefined,
    display: ellipsisLineCount ? '-webkit-box' : columnCount > 1 ? 'block' : undefined,
    WebkitBoxOrient: ellipsisLineCount ? 'vertical' : undefined,
    WebkitBoxPack: ellipsisLineCount ? element.verticalAlign === 'middle' ? 'center' : element.verticalAlign === 'bottom' ? 'end' : 'start' : undefined,
    WebkitLineClamp: ellipsisLineCount,
    columnCount: columnCount > 1 ? columnCount : undefined,
    columnGap: columnCount > 1 ? `${(element.textColumnSpacing ?? 0) * scale.value}px` : undefined,
    columnFill: columnCount > 1 ? 'auto' : undefined,
    padding: `${margins.top * scale.value}px ${margins.right * scale.value}px ${margins.bottom * scale.value}px ${margins.left * scale.value}px`,
    justifyContent: element.verticalAlign === 'middle' ? 'center' : element.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start',
    overflowX: element.textHorizontalOverflow === undefined ? undefined : element.textHorizontalOverflow === 'clip' ? 'clip' : 'visible',
    overflowY: ellipsisLineCount ? 'hidden' : element.textVerticalOverflow === undefined ? undefined : element.textVerticalOverflow === 'overflow' ? 'visible' : 'clip',
    whiteSpace: element.textWrap === 'none' ? 'nowrap' : undefined,
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textverticalvalues?view=openxml-3.0.1
    // https://www.w3.org/TR/css-writing-modes-3/#block-flow
    // wordArtVertRtl stacks text top-to-bottom and places earlier paragraphs to the right.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.vertical?view=openxml-3.0.0
    writingMode: verticalColumnWritingMode ?? (element.textOrientation === 'vert270' || mongolianVerticalText ? 'vertical-lr' : verticalText || eastAsianVerticalText || wordArtVerticalText || wordArtVerticalRtlText ? 'vertical-rl' : undefined),
    textOrientation: verticalText ? 'sideways' : eastAsianVerticalText || wordArtVerticalText || wordArtVerticalRtlText || mongolianVerticalText ? 'upright' : undefined,
    direction: element.textOrientation === 'vert270' || usesRightToLeftColumns(element) ? 'rtl' : undefined,
  }
}

function textBoundsStyle(element: PptxElement): CSSProperties {
  if (!element.textAnchorCenter) return { display: 'contents' }
  const centeredOffset = textAnchorCenterOffsets.get(`${props.slide.id}:${element.id}`)
  return {
    display: (element.textColumnCount ?? 1) > 1 ? 'block' : 'flex',
    flexDirection: (element.textColumnCount ?? 1) > 1 ? undefined : 'column',
    width: 'fit-content',
    maxWidth: '100%',
    marginLeft: 'auto',
    marginRight: 'auto',
    transform: centeredOffset ? `translateX(${centeredOffset}px)` : undefined,
  }
}

function effectiveTextFontScale(element: PptxElement): number {
  return normalAutoFitScales.get(`${props.slide.id}:${element.id}`) ?? element.textFontScale ?? 1
}

function tableCellAutoFitKey(elementId: string, rowIndex: number, cellIndex: number): string {
  return `${props.slide.id}:${elementId}:${rowIndex}:${cellIndex}`
}

function effectiveTableCellFontScale(element: PptxElement, rowIndex: number, cellIndex: number, cell: PptxTableCell): number {
  return normalAutoFitScales.get(tableCellAutoFitKey(element.id, rowIndex, cellIndex)) ?? cell.textFontScale ?? 1
}

function geometryStyle(element: PptxElement): CSSProperties {
  const geometry = element.geometry
  if (['ellipse', 'flowChartConnector', 'flowChartSummingJunction', 'flowChartOr'].includes(geometry)) return { borderRadius: '50%' }
  if (geometry === 'roundRect' || geometry === 'flowChartAlternateProcess') { const [x, y] = roundRectRadii(element); return { borderRadius: `${x / 10}% / ${y / 10}%` } }
  const clipPaths: Record<string, string> = {
  }
  return clipPaths[geometry] ? { clipPath: clipPaths[geometry] } : {}
}


type TextSegment = { text: string; script: ScriptFont; start: number; end: number; separateDecorations?: boolean }


// shortcut: SVG filters approximate text and irregular-shape inner shadows; calibrate against Office/WPS screenshots before claiming pixel parity.
function textInnerShadows(run: PptxElement['paragraphs'][number]['runs'][number]): PptxShadow[] {
  const shadows = run.shadows?.length ? run.shadows : run.shadow ? [run.shadow] : []
  return shadows.filter(shadow => shadow.inset)
}

function textInnerShadowFilterId(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number, filterKey: string): string {
  return `${customGradientScope}-text-inner-shadow-${element.id}-${filterKey}-${paragraphIndex}-${characterStart}-${characterEnd}`.replace(/[^a-zA-Z0-9_-]/g, '_')
}

function textInnerShadowFilterReference(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], paragraphIndex: number, characterStart: number, characterEnd: number, filterKey: string): string | undefined {
  return textInnerShadows(run).length ? `url(#${textInnerShadowFilterId(element, paragraphIndex, characterStart, characterEnd, filterKey)})` : undefined
}

function textInnerShadowOffset(element: PptxElement, shadow: PptxShadow, paragraphIndex: number, characterStart: number, characterEnd: number, shadowIndex: number): {dx: number; dy: number} {
  const [x, y] = animationShadowOffset(element, shadow, shadowIndex, paragraphIndex, characterStart, characterEnd)
  return {dx: x * scale.value, dy: y * scale.value}
}

function textInnerShadowColor(element: PptxElement, shadow: PptxShadow, paragraphIndex: number, characterStart: number, characterEnd: number): string {
  const color = animationColor(element, 'shadow.color', paragraphIndex, characterStart, characterEnd) || shadow.color
  return animatedRunColor(color, animationOpacity(element, 'shadow.opacity', paragraphIndex, characterStart, characterEnd)) || color
}

function runTextShadowStyle(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties['textShadow'] {
  const shadows = (run.shadows?.length ? run.shadows : run.shadow ? [run.shadow] : []).filter(shadow => !shadow.inset)
  const glows = run.glows?.length ? run.glows : run.glow ? [run.glow] : []
  const effects = shadows.map((shadow, shadowIndex) => {
    const shadowColor = animationColor(element, 'shadow.color', paragraphIndex, characterStart, characterEnd) || shadow.color
    return shadowCss(shadow, animatedRunColor(shadowColor, animationOpacity(element, 'shadow.opacity', paragraphIndex, characterStart, characterEnd)) || shadowColor, 1, { scaleX: 1, scaleY: 1 }, animationShadowOffset(element, shadow, shadowIndex, paragraphIndex, characterStart, characterEnd))
  })
  effects.push(...glows.map(glow => `0 0 ${glow.radius * scale.value}px ${glow.color}`))
  return effects.length ? effects.join(', ') : undefined
}

function cssBackgroundLayerCount(background: string): number {
  let depth = 0
  let layers = 1
  for (const character of background) {
    if (character === '(') depth++
    else if (character === ')') depth--
    else if (character === ',' && depth === 0) layers++
  }
  return layers
}

function runStyle(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], script: ScriptFont, fontScale: number, paragraphIndex: number, characterStart: number, characterEnd: number, separateDecorations = false, filterKey = `${paragraphIndex}-${characterStart}-${characterEnd}`): CSSProperties {
  const requestedFont = script === 'eastAsia' ? run.fontFamilyEastAsia : script === 'complexScript' ? run.fontFamilyComplexScript : run.fontFamily
  const fillOpacity = animationOpacity(element, 'fill.opacity', paragraphIndex, characterStart, characterEnd)
  const strokeOpacity = animationOpacity(element, 'stroke.opacity', paragraphIndex, characterStart, characterEnd)
  const textFill = run.fillGradient ? applyColorOpacity(run.fillGradient, fillOpacity) : run.fillPattern ? applyColorOpacity(run.fillPattern, fillOpacity) : undefined
  const highlightColor = run.highlightColor ? applyColorOpacity(run.highlightColor, fillOpacity) : undefined
  const highlightBackground = textFill && highlightColor ? `linear-gradient(${highlightColor}, ${highlightColor}) border-box, ${textFill}` : undefined
  const textFillClip = textFill && highlightColor
    ? `border-box, ${Array(cssBackgroundLayerCount(textFill)).fill('text').join(', ')}`
    : undefined
  const textDecorationLine = separateDecorations ? run.lineThrough ? 'line-through' : undefined : [run.underline ? 'underline' : undefined, run.lineThrough ? 'line-through' : undefined].filter(Boolean).join(' ')
  const textDecorationStyle = separateDecorations
    ? run.lineThroughStyle
    : run.underline && run.lineThrough
      ? (run.underlineStyle ?? (run.lineThroughStyle === 'double' ? 'solid' : run.lineThroughStyle))
      : run.underlineStyle ?? run.lineThroughStyle
  return {
    fontFamily: cssFontFamily(requestedFont || run.fontFamily),
    fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value}px`,
    letterSpacing: run.letterSpacingPt === undefined ? undefined : `${run.letterSpacingPt * fontScale * EMU_PER_POINT * scale.value}px`,
    verticalAlign: run.baselineOffsetEm === undefined ? undefined : `${run.baselineOffsetEm}em`,
    color: animatedRunColor(run.color, fillOpacity),
    '--pptx-text-fill': textFill,
    '--pptx-text-fill-color': animatedRunColor(run.color, fillOpacity),
    '--pptx-text-highlight': highlightColor,
    '--pptx-text-highlight-background': highlightBackground,
    '--pptx-text-fill-clip': textFillClip,
    WebkitTextStroke: run.outlineColor !== undefined && run.outlineWidth !== undefined ? `${run.outlineWidth * scale.value}px ${animatedRunColor(run.outlineColor, strokeOpacity)}` : undefined,
    fontWeight: run.bold ? 700 : undefined,
    fontStyle: run.italic ? 'italic' : undefined,
    ...(separateDecorations ? { '--pptx-underline-line': run.underline ? 'underline' : 'none' } : {}),
    textDecorationLine: textDecorationLine || undefined,
    textDecorationStyle,
    textDecorationColor: !separateDecorations && run.underline ? run.underlineColor : undefined,
    textDecorationThickness: !separateDecorations && run.underline && run.underlineThicknessEmu ? `${run.underlineThicknessEmu * scale.value}px` : undefined,
    textShadow: runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd),
    filter: textInnerShadowFilterReference(element, run, paragraphIndex, characterStart, characterEnd, filterKey),
  }
}

function textUnderlineStyle(run: PptxElement['paragraphs'][number]['runs'][number]): CSSProperties {
  return {
    textDecorationLine: 'var(--pptx-underline-line, none)',
    textDecorationStyle: run.underlineStyle,
    textDecorationColor: run.underlineColor,
    textDecorationThickness: run.underlineThicknessEmu ? `${run.underlineThicknessEmu * scale.value}px` : undefined,
  }
}

function runCharacterOffset(element: PptxElement, paragraphIndex: number, runIndex: number): number {
  let offset = 0
  for (let index = 0; index < paragraphIndex; index++) {
    const paragraph = element.paragraphs[index]
    if (paragraph) offset += paragraph.runs.reduce((sum, run) => sum + run.text.length, 0) + 1
  }
  const paragraph = element.paragraphs[paragraphIndex]
  if (paragraph) for (let index = 0; index < runIndex; index++) offset += paragraph.runs[index]?.text.length || 0
  return offset
}

function separateTextDecorations(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], paragraphIndex: number, characterStart: number, characterEnd: number): boolean {
  if (run.underline && run.lineThrough) return true
  return (animationsByTarget.value.get(element.id) || []).some((action) => {
    if (action.effect !== 'textStyle' || (action.textStyleProperty !== 'underline' && action.textStyleProperty !== 'lineThrough')) return false
    const iterationIndex = iterationIndexFor(action, characterStart, characterEnd)
    return (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
      && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
      && (!action.iteration || iterationIndex !== undefined && iterationIndex >= 0)
  })
}

function runSegments(run: PptxElement['paragraphs'][number]['runs'][number], element: PptxElement, paragraphIndex: number, runIndex: number): TextSegment[] {
  const runStart = runCharacterOffset(element, paragraphIndex, runIndex)
  const actions = animationsByTarget.value.get(element.id) || []
  const characterRanges = actions.flatMap(action => [
    ...(action.characterRange ? [action.characterRange] : []),
    ...(action.iteration?.ranges || []),
  ])
  if (!characterRanges.length) {
    const segments: TextSegment[] = []
    let segmentStart = 0
    let segmentScript: ScriptFont | undefined
    let previousScript: ScriptFont = 'latin'
    let offset = 0
    for (const character of run.text) {
      const script = characterScript(character, previousScript)
      if (segmentScript && script !== segmentScript) {
        segments.push({ text: run.text.slice(segmentStart, offset), script: segmentScript, start: runStart + segmentStart, end: runStart + offset })
        segmentStart = offset
      }
      segmentScript = script
      previousScript = script
      offset += character.length
    }
    if (segmentScript && offset > segmentStart) segments.push({ text: run.text.slice(segmentStart), script: segmentScript, start: runStart + segmentStart, end: runStart + offset })
    return segments.map(segment => ({ ...segment, separateDecorations: separateTextDecorations(element, run, paragraphIndex, segment.start, segment.end) }))
  }

  const characters: { text: string; script: ScriptFont; start: number; end: number }[] = []
  const cuts = new Set([0, run.text.length])
  let previousScript: ScriptFont = 'latin'
  let offset = 0
  for (const character of run.text) {
    const script = characterScript(character, previousScript)
    const end = offset + character.length
    if (characters.length && characters.at(-1)?.script !== script) cuts.add(offset)
    characters.push({ text: character, script, start: offset, end })
    previousScript = script
    offset = end
  }
  const boundaries = new Set<number>([0])
  for (const character of characters) boundaries.add(character.end)
  for (const range of characterRanges) {
    for (const boundary of [range.start, range.end]) {
      const local = boundary - runStart
      if (local > 0 && local < run.text.length && boundaries.has(local)) cuts.add(local)
    }
  }
  const sortedCuts = [...cuts].sort((left, right) => left - right)
  const segments: TextSegment[] = []
  let characterIndex = 0
  for (let index = 1; index < sortedCuts.length; index++) {
    const start = sortedCuts[index - 1]!
    const end = sortedCuts[index]!
    while (characters[characterIndex] && characters[characterIndex]!.end <= start) characterIndex++
    const first = characters[characterIndex]
    if (!first || end <= start) continue
    segments.push({ text: run.text.slice(start, end), script: first.script, start: runStart + start, end: runStart + end })
  }
  return segments.map(segment => ({ ...segment, separateDecorations: separateTextDecorations(element, run, paragraphIndex, segment.start, segment.end) }))
}

function paragraphBulletStyle(paragraph: PptxParagraph): CSSProperties {
  const size = paragraph.bullet?.size
  if (!size) return {}
  return { fontSize: size.unit === 'ratio' ? `${size.value * 100}%` : `${size.value * EMU_PER_POINT * scale.value}px` }
}

function cssSpacing(spacing: PptxParagraph['lineSpacing']): string | undefined {
  if (!spacing) return undefined
  return spacing.unit === 'emu' ? `${spacing.value * scale.value}px` : `${spacing.value}em`
}

function paragraphStyle(paragraph: PptxParagraph, fontScale = 1, lineSpaceReduction = 0, fallbackDirection?: CSSProperties['direction']): CSSProperties {
  const lineSpacing = paragraph.lineSpacing
  const fontSizePt = paragraph.runs.reduce((size, run) => Math.max(size, run.fontSizePt || 18), 0) || 18
  return {
    fontSize: `${fontSizePt * fontScale * EMU_PER_POINT * scale.value}px`,
    textAlign: paragraph.align,
    // Keep paragraph bidi direction separate from the multicol container's direction.
    // https://www.w3.org/TR/css-writing-modes-4/#direction
    direction: paragraph.direction ?? fallbackDirection,
    tabSize: paragraph.tabStopInterval === undefined ? undefined : `${paragraph.tabStopInterval * scale.value}px`,
    marginLeft: paragraph.marginLeft === undefined ? undefined : `${paragraph.marginLeft * scale.value}px`,
    marginRight: paragraph.marginRight === undefined ? undefined : `${paragraph.marginRight * scale.value}px`,
    textIndent: paragraph.indent === undefined ? undefined : `${paragraph.indent * scale.value}px`,
    marginTop: cssSpacing(paragraph.spaceBefore),
    marginBottom: cssSpacing(paragraph.spaceAfter),
    lineHeight: lineSpacing?.unit === 'ratio' ? lineSpacing.value * (1 - lineSpaceReduction) : cssSpacing(lineSpacing),
  }
}

function tableCellStyle(cell: PptxTableCell): CSSProperties {
  const border = (side: keyof PptxTableCell['borders']) => {
    const value = cell.borders[side]
    return value ? `${Math.max(value.width * scale.value, 1)}px solid ${value.color}` : undefined
  }
  return {
    background: cell.fill === 'transparent' ? undefined : cell.fill,
    padding: `${cell.margins.top * scale.value}px ${cell.margins.right * scale.value}px ${cell.margins.bottom * scale.value}px ${cell.margins.left * scale.value}px`,
    verticalAlign: cell.verticalAlign,
    borderTop: border('top'), borderRight: border('right'),
    borderBottom: border('bottom'), borderLeft: border('left'),
  }
}

function tableCellTextStyle(cell: PptxTableCell, rowHeight: number): CSSProperties {
  const columnCount = cell.textColumnCount ?? 1
  return {
    ...(columnCount > 1 || cell.textAutoFitDynamic ? {
      height: rowHeight > 0 ? `${Math.max(0, rowHeight - cell.margins.top - cell.margins.bottom) * scale.value}px` : '100%',
      boxSizing: 'border-box',
      minHeight: 0,
      overflow: 'hidden',
      ...(columnCount > 1 ? {
        columnCount,
        columnGap: `${(cell.textColumnSpacing ?? 0) * scale.value}px`,
        columnFill: 'auto',
        direction: cell.textColumnsRightToLeft ? 'rtl' : undefined,
      } : {}),
    } : {}),
    ...(cell.textAnchorCenter ? { width: 'fit-content', maxWidth: '100%', marginLeft: 'auto', marginRight: 'auto' } : {}),
  }
}

function tableStyle(element: PptxElement): CSSProperties {
  if (!element.renderMatrix) return {}
  const { scaleX, scaleY } = elementRenderSize(element)
  return {
    width: `${100 / scaleX}%`,
    height: `${100 / scaleY}%`,
    transform: `scale(${scaleX},${scaleY})`,
    transformOrigin: '0 0',
  }
}
</script>

<template>
  <div
    ref="stage"
    class="slide-frame"
    :class="{ 'is-thumbnail': thumbnail }"
    :style="frameStyle"
    :role="hasInteractiveMedia ? 'group' : 'img'"
    :aria-label="t('slideNumber', { number })"
  >
    <div class="slide-background-layer" :style="backgroundStyle" aria-hidden="true" />
    <div class="slide-content-layer">
    <div
      v-for="element in slide.elements"
      v-memo="elementMemoDependencies(element)"
      :key="element.id"
      :data-element-id="element.id"
      :data-element-name="element.name"
      :data-element-kind="element.kind"
      class="slide-element"
      :class="{
        'slide-shape': element.kind === 'shape',
        'shape-line': element.kind === 'line',
        'slide-placeholder': element.kind === 'placeholder',
        'is-animation-trigger': hasTrigger(element.id),
        'is-hyperlink': canActivateHyperlink(element.hyperlink),
      }"
      :role="canActivateHyperlink(element.hyperlink) ? 'link' : hasTrigger(element.id) ? 'button' : undefined"
      :tabindex="canActivateHyperlink(element.hyperlink) || hasTrigger(element.id) ? 0 : undefined"
      :aria-label="element.hyperlink?.tooltip || (hasTrigger(element.id) ? t('playAnimationFrom', { name: element.name }) : undefined)"
      :data-pptx-hyperlink="canActivateHyperlink(element.hyperlink) ? JSON.stringify(element.hyperlink) : undefined"
      :style="[elementStyle(element), geometryStyle(element), animationStyle(element)]"
      :title="element.hyperlink?.tooltip || element.name"
      @click="onElementClick($event, element.id, element.hyperlink)"
      @dblclick="onElementDoubleClick($event, element.id)"
      @mouseenter="onElementMouseEnter($event, element.id)"
      @mouseleave="onElementMouseLeave($event, element.id)"
      @keydown.enter.space.self.stop.prevent="onElementKeyActivate($event, element.id, element.hyperlink)"
    >
      <svg v-if="usesSvgOutline(element)" class="preset-outline" viewBox="0 0 1000 1000" preserveAspectRatio="none" :style="usesSvgPolygon(element) ? customGeometryStyle(element) : undefined" aria-hidden="true">
        <defs>
          <filter v-if="usesSvgPolygon(element) && customGeometryInnerShadows(element).length && animationOnOff(element, 'shadow.on')" :id="customGeometryInnerShadowFilterId(element)" v-bind="customGeometryInnerShadowFilterRegion(element)" filterUnits="objectBoundingBox" primitiveUnits="objectBoundingBox" color-interpolation-filters="sRGB">
            <template v-for="(shadow, shadowIndex) in customGeometryInnerShadows(element)" :key="shadowIndex">
              <feGaussianBlur in="SourceAlpha" :stdDeviation="customGeometryInnerShadowDeviation(element, shadow)" :result="'custom-inner-blur-' + shadowIndex" />
              <feOffset :in="'custom-inner-blur-' + shadowIndex" :dx="customGeometryInnerShadowOffset(element, shadow).dx" :dy="customGeometryInnerShadowOffset(element, shadow).dy" :result="'custom-inner-offset-' + shadowIndex" />
              <feComposite in="SourceAlpha" :in2="'custom-inner-offset-' + shadowIndex" operator="out" :result="'custom-inner-mask-' + shadowIndex" />
              <feFlood :flood-color="customGeometryInnerShadowColor(element, shadow)" :result="'custom-inner-color-' + shadowIndex" />
              <feComposite :in="'custom-inner-color-' + shadowIndex" :in2="'custom-inner-mask-' + shadowIndex" operator="in" :result="'custom-inner-shadow-' + shadowIndex" />
            </template>
            <feMerge>
              <feMergeNode v-for="(_, shadowIndex) in customGeometryInnerShadows(element)" :key="shadowIndex" :in="'custom-inner-shadow-' + shadowIndex" />
              <feMergeNode in="SourceGraphic" />
            </feMerge>
          </filter>
          <template v-for="gradient in customGeometryGradients(element).filter(item => usesSvgPolygon(element) || item.id === customGradientId(element, 'stroke'))" :key="gradient.id">
            <linearGradient v-if="gradient.kind === 'linear'" :id="gradient.id" gradientUnits="userSpaceOnUse" :x1="gradient.x1" :y1="gradient.y1" :x2="gradient.x2" :y2="gradient.y2"><stop v-for="(stop, index) in gradient.stops" :key="index" :offset="stop.offset" :stop-color="stop.color" :stop-opacity="stop.opacity" /></linearGradient>
            <radialGradient v-else :id="gradient.id" gradientUnits="userSpaceOnUse" :cx="gradient.cx" :cy="gradient.cy" :r="gradient.r" :fx="gradient.fx" :fy="gradient.fy" :fr="gradient.fr" :gradientTransform="gradient.gradientTransform" :color-interpolation="gradient.colorInterpolation"><stop v-for="(stop, index) in gradient.stops" :key="index" :offset="stop.offset" :stop-color="stop.color" :stop-opacity="stop.opacity" /></radialGradient>
          </template>
          <pattern v-for="mode in customGeometryPatternModes(element)" :key="mode || 'norm'" :id="customGeometryPatternId(element, mode)" patternUnits="userSpaceOnUse" width="1000" height="1000"><foreignObject width="1000" height="1000"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%" :style="{ background: applyColorOpacity(element.fill, 1, customPathBrightness(mode)) }" /></foreignObject></pattern>
          <pattern v-if="element.strokePattern" :id="customGeometryStrokePatternId(element)" patternUnits="userSpaceOnUse" width="1000" height="1000"><foreignObject width="1000" height="1000"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%" :style="{ background: element.strokePattern }" /></foreignObject></pattern>
        </defs>
        <path :d="presetOutlinePath(element)" :fill="usesSvgPolygon(element) ? customGeometryFill(element, { d: '', fill: true, stroke: true }) : 'none'" :fill-opacity="usesSvgPolygon(element) ? customGeometryFillOpacity(element, { d: '', fill: true, stroke: true }) : undefined" :stroke="presetOutlineDetailPaths(element).length ? 'none' : customGeometryStroke(element, { d: '', fill: false, stroke: true })" :stroke-width="customGeometryStrokeWidth(element)" :style="customGeometryStrokeStyle(element)" vector-effect="non-scaling-stroke" />
        <path v-for="(detailPath, index) in presetOutlineDetailPaths(element)" :key="index" :d="detailPath" fill="none" :stroke="customGeometryStroke(element, { d: detailPath, fill: false, stroke: true })" :stroke-width="customGeometryStrokeWidth(element)" :style="customGeometryStrokeStyle(element)" vector-effect="non-scaling-stroke" />
      </svg>
      <svg v-if="element.kind === 'shape' && element.customPaths?.length" class="custom-geometry" viewBox="0 0 1000 1000" preserveAspectRatio="none" :style="customGeometryStyle(element)" aria-hidden="true">
        <defs>
          <filter v-if="customGeometryInnerShadows(element).length && animationOnOff(element, 'shadow.on')" :id="customGeometryInnerShadowFilterId(element)" v-bind="customGeometryInnerShadowFilterRegion(element)" filterUnits="objectBoundingBox" primitiveUnits="objectBoundingBox" color-interpolation-filters="sRGB">
            <template v-for="(shadow, shadowIndex) in customGeometryInnerShadows(element)" :key="shadowIndex">
              <feGaussianBlur in="SourceAlpha" :stdDeviation="customGeometryInnerShadowDeviation(element, shadow)" :result="'custom-inner-blur-' + shadowIndex" />
              <feOffset :in="'custom-inner-blur-' + shadowIndex" :dx="customGeometryInnerShadowOffset(element, shadow).dx" :dy="customGeometryInnerShadowOffset(element, shadow).dy" :result="'custom-inner-offset-' + shadowIndex" />
              <feComposite in="SourceAlpha" :in2="'custom-inner-offset-' + shadowIndex" operator="out" :result="'custom-inner-mask-' + shadowIndex" />
              <feFlood :flood-color="customGeometryInnerShadowColor(element, shadow)" :result="'custom-inner-color-' + shadowIndex" />
              <feComposite :in="'custom-inner-color-' + shadowIndex" :in2="'custom-inner-mask-' + shadowIndex" operator="in" :result="'custom-inner-shadow-' + shadowIndex" />
            </template>
            <feMerge>
              <feMergeNode v-for="(_, shadowIndex) in customGeometryInnerShadows(element)" :key="shadowIndex" :in="'custom-inner-shadow-' + shadowIndex" />
              <feMergeNode in="SourceGraphic" />
            </feMerge>
          </filter>
          <template v-for="gradient in customGeometryGradients(element)" :key="gradient.id">
            <linearGradient v-if="gradient.kind === 'linear'" :id="gradient.id" gradientUnits="userSpaceOnUse" :x1="gradient.x1" :y1="gradient.y1" :x2="gradient.x2" :y2="gradient.y2"><stop v-for="(stop, stopIndex) in gradient.stops" :key="stopIndex" :offset="stop.offset" :stop-color="stop.color" :stop-opacity="stop.opacity" /></linearGradient>
            <radialGradient v-else :id="gradient.id" gradientUnits="userSpaceOnUse" :cx="gradient.cx" :cy="gradient.cy" :r="gradient.r" :fx="gradient.fx" :fy="gradient.fy" :fr="gradient.fr" :gradientTransform="gradient.gradientTransform" :color-interpolation="gradient.colorInterpolation"><stop v-for="(stop, stopIndex) in gradient.stops" :key="stopIndex" :offset="stop.offset" :stop-color="stop.color" :stop-opacity="stop.opacity" /></radialGradient>
          </template>
          <pattern v-for="mode in customGeometryPatternModes(element)" :key="mode || 'norm'" :id="customGeometryPatternId(element, mode)" patternUnits="userSpaceOnUse" width="1000" height="1000"><foreignObject width="1000" height="1000"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%" :style="{ background: applyColorOpacity(element.fill, 1, customPathBrightness(mode)) }" /></foreignObject></pattern>
          <pattern v-if="element.strokePattern" :id="customGeometryStrokePatternId(element)" patternUnits="userSpaceOnUse" width="1000" height="1000"><foreignObject width="1000" height="1000"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%" :style="{ background: element.strokePattern }" /></foreignObject></pattern>
          <template v-for="(path, pathIndex) in element.customPaths" :key="`focus-${pathIndex}`"><clipPath v-if="customGeometryFillFocus(element, path)" :id="customGradientFocusClipId(element, pathIndex)" clipPathUnits="userSpaceOnUse"><path :d="path.d" /></clipPath></template>
        </defs>
        <template v-for="(path, pathIndex) in element.customPaths" :key="pathIndex">
          <path :d="path.d" :fill="customGeometryFill(element, path)" :fill-opacity="customGeometryFillOpacity(element, path)" :stroke="customGeometryStroke(element, path)" :stroke-width="customGeometryStrokeWidth(element)" :style="customGeometryStrokeStyle(element)" vector-effect="non-scaling-stroke" />
          <template v-for="focus in [customGeometryFillFocus(element, path)]" :key="`focus-${pathIndex}`">
            <ellipse v-if="focus?.kind === 'ellipse'" :cx="focus.cx" :cy="focus.cy" :rx="focus.rx" :ry="focus.ry" :fill="focus.color" :clip-path="`url(#${customGradientFocusClipId(element, pathIndex)})`" />
            <rect v-else-if="focus" :x="focus.x" :y="focus.y" :width="focus.width" :height="focus.height" :fill="focus.color" :clip-path="`url(#${customGradientFocusClipId(element, pathIndex)})`" />
          </template>
        </template>
      </svg>
      <div v-if="element.kind === 'video' && element.mediaUrl && !thumbnail" class="image-viewport">
        <video class="slide-video" :data-media-id="element.id" :data-media-slide-count="mediaSlideCounts.get(element.id) || undefined" :src="element.mediaUrl" :poster="element.imageUrl" controls playsinline preload="metadata" :aria-label="element.name" :style="imageStyle(element)" @loadedmetadata="updateMediaPlayback($event, element)" @seeking="updateMediaPlayback($event, element)" @timeupdate="updateMediaPlayback($event, element)" @play="updateMediaOnPlay($event, element)" @pause="updateMediaFadeForEvent($event, element)" @ended="finishTimedMedia(element, true, true)" @click.stop @pointerdown.stop />
        <span v-if="invalidMediaTrimIds.includes(element.id)" class="media-warning" role="status">{{ t('trimRangeWarning') }}</span>
        <span v-if="invalidMediaFadeIds.includes(element.id)" class="media-warning" role="status">{{ t('mediaFadeWarning') }}</span>
        <span v-if="blockedMediaPlaybackIds.includes(element.id)" class="media-warning" role="status">{{ t('automaticPlaybackBlocked') }}</span>
      </div>
      <div v-else-if="element.kind === 'audio' && element.mediaUrl && !thumbnail" class="audio-viewport">
        <audio class="slide-audio" :data-media-id="element.id" :data-media-slide-count="mediaSlideCounts.get(element.id) || undefined" :src="element.mediaUrl" controls preload="metadata" :aria-label="element.name" @loadedmetadata="updateMediaPlayback($event, element)" @seeking="updateMediaPlayback($event, element)" @timeupdate="updateMediaPlayback($event, element)" @play="updateMediaOnPlay($event, element)" @pause="updateMediaFadeForEvent($event, element)" @volumechange="updateMediaFadeForEvent($event, element)" @ended="finishTimedMedia(element, true, true)" @click.stop @pointerdown.stop />
        <span v-if="invalidMediaTrimIds.includes(element.id)" class="media-warning" role="status">{{ t('trimRangeWarning') }}</span>
        <span v-if="invalidMediaFadeIds.includes(element.id)" class="media-warning" role="status">{{ t('mediaFadeWarning') }}</span>
        <span v-if="blockedMediaPlaybackIds.includes(element.id)" class="media-warning" role="status">{{ t('automaticPlaybackBlocked') }}</span>
      </div>
      <svg v-else-if="element.kind === 'chart' && element.chart" class="slide-chart" :data-chart-type="element.chart.type" :data-radar-style="element.chart.radarStyle" viewBox="0 0 1000 600" preserveAspectRatio="none" role="img" :aria-label="element.chart.title || element.name">
        <title>{{ element.chart.title || element.name }}</title>
        <defs>
          <clipPath :id="`chart-plot-${element.id}`"><polygon :points="`${chartPlot(element.id).left},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top + chartPlot(element.id).height} ${chartPlot(element.id).left},${chartPlot(element.id).top + chartPlot(element.id).height}`" /></clipPath>
          <clipPath v-if="chartVisuals.get(element.id)?.legendClip" :id="chartVisuals.get(element.id)?.legendClip?.id"><rect :x="chartVisuals.get(element.id)?.legendClip?.x" :y="chartVisuals.get(element.id)?.legendClip?.y" :width="chartVisuals.get(element.id)?.legendClip?.width" :height="chartVisuals.get(element.id)?.legendClip?.height" /></clipPath>
          <pattern v-for="pattern in chartPatternDefinitions.get(element.id)" :key="pattern.id" :id="pattern.id" patternUnits="userSpaceOnUse" width="1000" height="600"><foreignObject width="1000" height="600"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%" :style="{ background: pattern.fill }" /></foreignObject></pattern>
        </defs>
        <text v-if="element.chart.title" x="500" y="42" text-anchor="middle" class="chart-title">{{ element.chart.title }}</text>
        <g v-if="element.chart.type === 'pie' || element.chart.type === 'doughnut'">
          <path v-for="(slice, index) in chartVisuals.get(element.id)?.slices" :key="index" :d="slice.d" :fill="chartPaint(element.id, slice.color)" :fill-rule="element.chart.type === 'doughnut' ? 'evenodd' : 'nonzero'" :stroke="chartPaint(element.id, slice.stroke)" :stroke-width="slice.strokeWidth ?? 2" :stroke-dasharray="slice.strokeDasharray" :stroke-linecap="slice.strokeLinecap" />
        </g>
        <g v-else>
          <template v-if="element.chart.type === 'radar'"><path v-for="(grid, index) in chartVisuals.get(element.id)?.radarGrid" :key="`radar-grid-${index}`" class="chart-major-gridline chart-radar-gridline" :d="grid.d" fill="none" stroke="#d1d5db" stroke-width="1" /><line v-for="(spoke, index) in chartVisuals.get(element.id)?.radarSpokes" :key="`radar-spoke-${index}`" class="chart-major-gridline chart-radar-spoke" v-bind="spoke" stroke="#d1d5db" stroke-width="1" /></template>
          <template v-else>
            <line v-if="!element.chart.categoryAxisDeleted && element.chart.categoryAxisMajorGridlines" v-for="(grid, index) in chartVisuals.get(element.id)?.categoryGridlines" :key="`category-grid-${index}`" class="chart-major-gridline chart-category-major-gridline" data-axis="category" v-bind="grid" stroke="#d1d5db" stroke-width="1" />
            <template v-if="!element.chart.valueAxis?.deleted && element.chart.valueAxis?.majorGridlines"><line v-for="(tick, index) in chartVisuals.get(element.id)?.ticks" :key="`grid-${index}`" class="chart-major-gridline" :x1="element.chart.direction === 'horizontal' ? tick.x : chartPlot(element.id).left" :x2="element.chart.direction === 'horizontal' ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width" :y1="element.chart.direction === 'horizontal' ? chartPlot(element.id).top : tick.y" :y2="element.chart.direction === 'horizontal' ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y" stroke="#d1d5db" stroke-width="1" /></template>
            <template v-for="axis in chartVisuals.get(element.id)?.secondaryValueAxes" :key="`secondary-grid-${axis.id}`"><line v-if="!axis.deleted && axis.majorGridlines" v-for="(tick, index) in axis.ticks" :key="index" class="chart-major-gridline chart-secondary-gridline" :x1="axis.horizontal ? tick.x : chartPlot(element.id).left" :x2="axis.horizontal ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width" :y1="axis.horizontal ? chartPlot(element.id).top : tick.y" :y2="axis.horizontal ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y" stroke="#d1d5db" stroke-width="1" /></template>
            <template v-if="['scatter', 'bubble'].includes(element.chart.type) && !element.chart.xAxis?.deleted && element.chart.xAxis?.majorGridlines"><line v-for="(tick, index) in chartVisuals.get(element.id)?.xTicks" :key="`x-grid-${index}`" class="chart-major-gridline chart-x-gridline" :x1="tick.x" :x2="tick.x" :y1="chartPlot(element.id).top" :y2="chartPlot(element.id).top + chartPlot(element.id).height" stroke="#d1d5db" stroke-width="1" /></template>
          </template>
          <template v-for="(tick, index) in chartVisuals.get(element.id)?.ticks" :key="`tick-${index}`"><text v-if="tick.showLabel" :x="tick.labelX" :y="tick.labelY" :text-anchor="tick.labelAnchor" :style="{ fill: tick.color }" class="chart-label chart-value-axis-label">{{ tick.label }}</text></template>
          <template v-for="axis in chartVisuals.get(element.id)?.secondaryValueAxes" :key="`secondary-tick-${axis.id}`"><template v-for="(tick, index) in axis.ticks" :key="index"><text v-if="tick.showLabel" :data-axis-id="axis.id" :x="tick.labelX" :y="tick.labelY" :text-anchor="tick.labelAnchor" :style="{ fill: tick.color }" class="chart-label chart-value-axis-label chart-secondary-value-axis-label">{{ tick.label }}</text></template></template>
          <template v-for="(tick, index) in chartVisuals.get(element.id)?.xTicks" :key="`x-tick-${index}`"><text v-if="tick.showLabel" :x="tick.labelX" :y="tick.labelY" :text-anchor="tick.labelAnchor" :style="{ fill: tick.color }" class="chart-label chart-x-value-axis-label">{{ tick.label }}</text></template>
          <line v-for="(tick, index) in chartVisuals.get(element.id)?.axisTicks" :key="`axis-tick-${index}`" :class="tick.minor ? 'chart-axis-minor-tick' : 'chart-axis-tick'" :data-axis="tick.axis" :x1="tick.x1" :y1="tick.y1" :x2="tick.x2" :y2="tick.y2" stroke="#64748b" :stroke-width="tick.minor ? 1 : 1.5" />
          <g :clip-path="`url(#chart-plot-${element.id})`">
            <line v-if="element.chart.type !== 'radar' && (['scatter', 'bubble'].includes(element.chart.type) ? !element.chart.xAxis?.deleted : !element.chart.categoryAxisDeleted) && chartVisuals.get(element.id)?.baseline" class="chart-category-axis-line" :class="{ 'chart-x-axis-line': ['scatter', 'bubble'].includes(element.chart.type) }" :data-axis="['scatter', 'bubble'].includes(element.chart.type) ? 'xValue' : 'category'" :x1="chartVisuals.get(element.id)?.baseline.x1" :y1="chartVisuals.get(element.id)?.baseline.y1" :x2="chartVisuals.get(element.id)?.baseline.x2" :y2="chartVisuals.get(element.id)?.baseline.y2" stroke="#64748b" stroke-width="2" />
            <line v-if="!element.chart.valueAxis?.deleted && chartVisuals.get(element.id)?.valueAxisLine" class="chart-value-axis-line" :x1="chartVisuals.get(element.id)?.valueAxisLine?.x1" :y1="chartVisuals.get(element.id)?.valueAxisLine?.y1" :x2="chartVisuals.get(element.id)?.valueAxisLine?.x2" :y2="chartVisuals.get(element.id)?.valueAxisLine?.y2" stroke="#64748b" stroke-width="2" />
            <template v-for="axis in chartVisuals.get(element.id)?.secondaryValueAxes" :key="`secondary-axis-${axis.id}`"><line v-if="!axis.deleted" class="chart-value-axis-line chart-secondary-value-axis-line" :data-axis-id="axis.id" :x1="axis.line.x1" :y1="axis.line.y1" :x2="axis.line.x2" :y2="axis.line.y2" stroke="#64748b" stroke-width="2" /></template>
            <path v-for="(area, index) in chartVisuals.get(element.id)?.areas" :key="`area-${index}`" class="chart-area" :d="area.d" :fill="chartPaint(element.id, area.color)" :fill-opacity="area.opacity ?? 1" :stroke="chartPaint(element.id, area.stroke)" :stroke-width="area.strokeWidth ?? 1" :stroke-dasharray="area.strokeDasharray" :stroke-linecap="area.strokeLinecap" />
            <rect v-for="(bar, index) in chartVisuals.get(element.id)?.bars" :key="`bar-${index}`" class="chart-bar" :x="bar.x" :y="bar.y" :width="bar.width" :height="bar.height" :fill="chartPaint(element.id, bar.color)" :stroke="chartPaint(element.id, bar.stroke)" :stroke-width="bar.strokeWidth ?? 1" :stroke-dasharray="bar.strokeDasharray" :stroke-linecap="bar.strokeLinecap" />
            <path v-for="(line, index) in chartVisuals.get(element.id)?.lines" :key="`line-${index}`" class="chart-series-line" :d="line.d" fill="none" :stroke="chartPaint(element.id, line.color)" :stroke-width="line.width ?? 4" :stroke-dasharray="line.strokeDasharray" :stroke-linecap="line.strokeLinecap ?? 'round'" />
            <template v-for="(marker, index) in chartVisuals.get(element.id)?.markers" :key="`marker-${index}`">
              <circle v-if="!chartMarkerPath(marker)" class="chart-line-marker" :class="{ 'chart-bubble-marker': element.chart.type === 'bubble' }" :data-marker-shape="marker.symbol ?? 'circle'" :data-x="marker.x" :data-y="marker.y" :data-radius="marker.radius ?? 5" :cx="marker.x" :cy="marker.y" :r="marker.symbol === 'dot' ? (marker.radius ?? 5) * 0.45 : marker.radius ?? 5" :fill="chartPaint(element.id, marker.color)" :stroke="chartPaint(element.id, marker.stroke)" :stroke-width="marker.strokeWidth ?? 1" :stroke-dasharray="marker.strokeDasharray" :stroke-linecap="marker.strokeLinecap" />
              <path v-else class="chart-line-marker" :data-marker-shape="marker.symbol" :data-x="marker.x" :data-y="marker.y" :data-radius="marker.radius ?? 5" :d="chartMarkerPath(marker)" :fill="chartPaint(element.id, marker.color)" :stroke="chartPaint(element.id, marker.stroke)" :stroke-width="marker.strokeWidth ?? 1" :stroke-dasharray="marker.strokeDasharray" :stroke-linecap="marker.strokeLinecap" />
            </template>
          </g>
          <text v-for="(label, index) in chartVisuals.get(element.id)?.categories" :key="`category-${index}`" :x="label.x" :y="label.y" :text-anchor="label.anchor" :style="{ fill: label.color }" class="chart-label chart-category-label">{{ label.text }}</text>
        </g>
        <g v-for="(label, index) in chartVisuals.get(element.id)?.dataLabels" :key="`data-label-${index}`">
          <path v-if="label.leaderLine" :d="label.leaderLine.d" :stroke="label.leaderLine.color" fill="none" stroke-width="1" class="chart-data-label-leader-line" />
          <g v-if="label.legendKey" :data-key-id="chartDataLabelKeyId(element.id, index)" :fill="chartPaint(element.id, label.legendKey.color)" class="chart-data-label-key">
            <rect v-if="!label.legendKey.markerChart" :x="chartDataLabelKeyPositions.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18" :y="label.y - 12" width="12" height="12" class="chart-data-label-key-swatch" />
            <path v-if="label.legendKey.line" :d="`M ${chartDataLabelKeyPositions.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18} ${label.y - 6} H ${(chartDataLabelKeyPositions.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 12}`" fill="none" :stroke="chartPaint(element.id, label.legendKey.color)" stroke-width="2" stroke-linecap="round" class="chart-data-label-key-line" />
            <template v-if="label.legendKey.symbol">
              <circle v-if="!chartMarkerPath({x: (chartDataLabelKeyPositions.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 6, y: label.y - 6, color: label.legendKey.color, radius: 4, symbol: label.legendKey.symbol})" :cx="(chartDataLabelKeyPositions.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 6" :cy="label.y - 6" :r="label.legendKey.symbol === 'dot' ? 1.8 : 4" class="chart-data-label-key-marker" :data-marker-shape="label.legendKey.symbol" />
              <path v-else :d="chartMarkerPath({x: (chartDataLabelKeyPositions.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 6, y: label.y - 6, color: label.legendKey.color, radius: 4, symbol: label.legendKey.symbol})" class="chart-data-label-key-marker" :data-marker-shape="label.legendKey.symbol" />
            </template>
          </g>
          <text :x="label.x" :y="label.y" :text-anchor="label.anchor" :style="{ fill: label.style?.color, fontFamily: label.style?.fontFamily ? cssFontFamily(label.style.fontFamily) : undefined, fontSize: label.style?.fontSizePt === undefined ? undefined : `${label.style.fontSizePt * 4 / 3}px`, fontWeight: label.style?.bold === undefined ? undefined : label.style.bold ? 700 : 400, fontStyle: label.style?.italic === undefined ? undefined : label.style.italic ? 'italic' : 'normal', letterSpacing: label.style?.letterSpacingPt === undefined ? undefined : `${label.style.letterSpacingPt * 4 / 3}px`, textDecorationLine: label.style?.underline === undefined && label.style?.lineThrough === undefined ? undefined : [label.style?.underline ? 'underline' : undefined, label.style?.lineThrough ? 'line-through' : undefined].filter(Boolean).join(' ') || 'none', textDecorationStyle: label.style?.underline ? label.style.underlineStyle : label.style?.lineThrough ? label.style.lineThroughStyle : undefined, textDecorationColor: label.style?.underline ? label.style.underlineColor : undefined, textDecorationThickness: label.style?.underline && label.style.underlineThicknessEmu !== undefined ? `${label.style.underlineThicknessEmu * 4 / (3 * EMU_PER_POINT)}px` : undefined }" class="chart-data-label"><tspan v-for="(line, lineIndex) in label.lines" :key="lineIndex" :x="label.x" :y="label.y + (lineIndex - (label.lines.length - 1) / 2) * 22" :dy="label.style?.baselineOffsetEm === undefined ? undefined : `${-label.style.baselineOffsetEm}em`">{{ line }}</tspan></text>
        </g>
        <g class="chart-legend-group" :clip-path="chartVisuals.get(element.id)?.legendClip ? `url(#${chartVisuals.get(element.id)?.legendClip?.id})` : undefined">
          <g v-for="(entry, index) in chartVisuals.get(element.id)?.legend" :key="`legend-${index}`">
            <g :fill="chartPaint(element.id, entry.color)" class="chart-legend-key" :data-symbol="entry.symbol">
              <rect v-if="!entry.markerChart" :x="entry.x" :y="entry.y - 12" width="12" height="12" />
              <path v-if="entry.line" :d="`M ${entry.x} ${entry.y - 6} H ${entry.x + 12}`" fill="none" :stroke="chartPaint(element.id, entry.color)" stroke-width="2" stroke-linecap="round" class="chart-legend-key-line" />
              <template v-if="entry.symbol">
                <circle v-if="!chartMarkerPath({x: entry.x + 6, y: entry.y - 6, color: entry.color, radius: 4, symbol: entry.symbol})" :cx="entry.x + 6" :cy="entry.y - 6" :r="entry.symbol === 'dot' ? 1.8 : 4" class="chart-legend-key-marker" :data-marker-shape="entry.symbol" />
                <path v-else :d="chartMarkerPath({x: entry.x + 6, y: entry.y - 6, color: entry.color, radius: 4, symbol: entry.symbol})" class="chart-legend-key-marker" :data-marker-shape="entry.symbol" />
              </template>
            </g>
            <text :x="entry.x + 18" :y="entry.y" class="chart-label chart-legend-label">{{ entry.text }}</text>
          </g>
        </g>
      </svg>
      <div v-else-if="(element.kind === 'picture' || element.kind === 'video') && element.imageUrl" class="image-viewport" aria-hidden="true">
        <svg v-if="element.kind === 'picture' && (element.imageAlphaInverse !== undefined || element.imageEffectDagAlphaInverse !== undefined)" class="image-tone-defs image-alpha-inverse-defs" aria-hidden="true" width="0" height="0">
          <defs>
            <template v-for="(alphaInverse, index) in [element.imageAlphaInverse === undefined ? undefined : {invertAlpha: true, color: element.imageAlphaInverse ?? undefined}, element.imageEffectDagAlphaInverse]" :key="index">
              <filter v-if="alphaInverse !== undefined" :id="imageAlphaInverseFilterId(element, index === 1)" class="image-alpha-inverse-filter" color-interpolation-filters="sRGB">
                <feComponentTransfer>
                  <feFuncR v-if="alphaInverse.color" type="linear" slope="0" :intercept="alphaInverse.color[0]" />
                  <feFuncG v-if="alphaInverse.color" type="linear" slope="0" :intercept="alphaInverse.color[1]" />
                  <feFuncB v-if="alphaInverse.color" type="linear" slope="0" :intercept="alphaInverse.color[2]" />
                  <feFuncA v-if="alphaInverse.invertAlpha" type="table" tableValues="1 0" />
                </feComponentTransfer>
              </filter>
            </template>
          </defs>
        </svg>
        <svg v-for="tone in (imageToneEnabled(element) ? [imageTone(element)] : [])" :key="element.id" class="image-tone-defs" aria-hidden="true" width="0" height="0">
          <defs>
            <filter :id="imageToneFilterId(element)" class="image-tone-filter" color-interpolation-filters="sRGB">
              <feComponentTransfer>
                <feFuncR class="image-tone-gain" type="linear" :slope="tone.gain" :intercept="(1 - tone.gain) / 2" />
                <feFuncG class="image-tone-gain" type="linear" :slope="tone.gain" :intercept="(1 - tone.gain) / 2" />
                <feFuncB class="image-tone-gain" type="linear" :slope="tone.gain" :intercept="(1 - tone.gain) / 2" />
              </feComponentTransfer>
              <feComponentTransfer>
                <feFuncR class="image-tone-blacklevel" type="linear" slope="1" :intercept="tone.blacklevel" />
                <feFuncG class="image-tone-blacklevel" type="linear" slope="1" :intercept="tone.blacklevel" />
                <feFuncB class="image-tone-blacklevel" type="linear" slope="1" :intercept="tone.blacklevel" />
              </feComponentTransfer>
              <feComponentTransfer>
                <feFuncR class="image-tone-gamma" type="gamma" amplitude="1" :exponent="1 / tone.gamma" offset="0" />
                <feFuncG class="image-tone-gamma" type="gamma" amplitude="1" :exponent="1 / tone.gamma" offset="0" />
                <feFuncB class="image-tone-gamma" type="gamma" amplitude="1" :exponent="1 / tone.gamma" offset="0" />
              </feComponentTransfer>
            </filter>
          </defs>
        </svg>
        <img class="slide-image" :src="element.imageUrl" :alt="element.name" :style="imageStyle(element)">
        <div v-if="pictureInnerShadowStyle(element)" class="picture-inner-shadow" :style="pictureInnerShadowStyle(element)" aria-hidden="true"></div>
      </div>
      <table v-else-if="element.kind === 'table' && element.table" class="slide-table" :style="tableStyle(element)">
        <colgroup>
          <col v-for="(column, columnIndex) in element.table.columns" :key="columnIndex" :style="{ width: `${column / Math.max(element.table.columns.reduce((sum, width) => sum + width, 0), 1) * 100}%` }">
        </colgroup>
        <tbody>
          <tr v-for="(row, rowIndex) in element.table.rows" :key="rowIndex" :style="{ height: `${row.height / Math.max(element.table.rows.reduce((sum, item) => sum + item.height, 0), 1) * 100}%` }">
            <template v-for="(cell, cellIndex) in row.cells" :key="cellIndex">
              <td v-if="!cell.hidden" :colspan="cell.colSpan" :rowspan="cell.rowSpan" :data-autofit="cell.textAutoFitDynamic ? 'normal' : undefined" :data-row-index="rowIndex" :data-cell-index="cellIndex" :style="tableCellStyle(cell)">
                <div class="table-cell-text" :style="tableCellTextStyle(cell, element.table.rows.slice(rowIndex, rowIndex + cell.rowSpan).reduce((height, item) => height + item.height, 0))">
                  <div v-for="(entry, paragraphIndex) in paragraphEntries(cell.paragraphs)" :key="paragraphIndex" class="text-paragraph" :style="paragraphStyle(entry.paragraph, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), cell.textLineSpacingReduction, cell.textColumnsRightToLeft && (cell.textColumnCount ?? 1) > 1 ? 'ltr' : undefined)">
                    <img v-if="entry.paragraph.bullet?.kind === 'image' && entry.paragraph.bullet.url" class="paragraph-picture-bullet" :src="entry.paragraph.bullet.url" :style="paragraphBulletStyle(entry.paragraph)" alt="" aria-hidden="true">
                    <span v-else-if="entry.marker" class="paragraph-marker" :style="paragraphBulletStyle(entry.paragraph)" aria-hidden="true">{{ entry.marker }}</span>
                    <template v-for="(run, runIndex) in entry.paragraph.runs" :key="runIndex">
                      <component
                        v-for="(segment, segmentIndex) in runSegments(run, element, paragraphIndex, runIndex)"
                        :key="segmentIndex"
                        :is="canActivateHyperlink(run.hyperlink) ? 'a' : 'span'"
                        :href="canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : undefined"
                        :title="run.hyperlink?.tooltip"
                        :data-pptx-hyperlink="canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : undefined"
                        :class="['text-run', { 'text-run-hyperlink': canActivateHyperlink(run.hyperlink), 'text-run-break': segment.text === '\n', 'text-run-fill': !segment.separateDecorations && runHasTextFill(element, run, paragraphIndex, segment.start, segment.end), 'text-run-highlight': Boolean(run.highlightColor) }]"
                        :data-char-start="segment.start"
                        :data-char-end="segment.end"
                :style="[runStyle(element, run, segment.script, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), paragraphIndex, segment.start, segment.end, segment.separateDecorations, 'table-' + rowIndex + '-' + cellIndex + '-' + runIndex), textColorAnimationStyle(element, paragraphIndex, segment.start, segment.end, run.color), textRangeOpacityAnimationStyle(element, paragraphIndex, segment.start, segment.end), characterAnimationStyle(element, paragraphIndex, segment.start, segment.end), textRunAnimationStyle(element, paragraphIndex, segment.start, segment.end, run, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), segment.script, segment.separateDecorations, 'table-' + rowIndex + '-' + cellIndex + '-' + runIndex), textTransformAnimationStyle(element, paragraphIndex, segment.start, segment.end)]"
                      >
                        <svg v-if="textInnerShadows(run).length" class="text-run-filter-defs" width="0" height="0" aria-hidden="true" focusable="false">
                          <defs>
                            <filter :id="textInnerShadowFilterId(element, paragraphIndex, segment.start, segment.end, 'table-' + rowIndex + '-' + cellIndex + '-' + runIndex)" x="-10%" y="-10%" width="120%" height="120%" filterUnits="objectBoundingBox" primitiveUnits="userSpaceOnUse" color-interpolation-filters="sRGB">
                              <template v-for="(shadow, shadowIndex) in textInnerShadows(run)" :key="shadowIndex">
                                <feGaussianBlur in="SourceAlpha" :stdDeviation="Math.max(0, shadow.blur * scale / 2)" :result="'inner-blur-' + shadowIndex" />
                                <feOffset :in="'inner-blur-' + shadowIndex" :dx="textInnerShadowOffset(element, shadow, paragraphIndex, segment.start, segment.end, shadowIndex).dx" :dy="textInnerShadowOffset(element, shadow, paragraphIndex, segment.start, segment.end, shadowIndex).dy" :result="'inner-offset-' + shadowIndex" />
                                <feComposite in="SourceAlpha" :in2="'inner-offset-' + shadowIndex" operator="out" :result="'inner-mask-' + shadowIndex" />
                                <feFlood :flood-color="textInnerShadowColor(element, shadow, paragraphIndex, segment.start, segment.end)" :result="'inner-color-' + shadowIndex" />
                                <feComposite :in="'inner-color-' + shadowIndex" :in2="'inner-mask-' + shadowIndex" operator="in" :result="'inner-shadow-' + shadowIndex" />
                              </template>
                              <feMerge>
                                <feMergeNode v-for="(_, shadowIndex) in textInnerShadows(run)" :key="shadowIndex" :in="'inner-shadow-' + shadowIndex" />
                                <feMergeNode in="SourceGraphic" />
                              </feMerge>
                            </filter>
                          </defs>
                        </svg>
                        <span v-if="segment.separateDecorations" class="text-run-underline" :class="{ 'text-run-fill': runHasTextFill(element, run, paragraphIndex, segment.start, segment.end), 'text-run-highlight': Boolean(run.highlightColor) }" :style="textUnderlineStyle(run)">{{ segment.text }}</span>
                        <template v-else>{{ segment.text }}</template>
                      </component>
                    </template>
                  </div>
                </div>
              </td>
            </template>
          </tr>
        </tbody>
      </table>
      <span v-else-if="element.kind === 'placeholder'">{{ t(element.mediaType ? 'mediaUnavailable' : 'imageUnavailable') }}</span>
      <div v-if="hasRenderableText(element)" class="text-frame" :class="{ 'text-frame-no-wrap': element.textWrap === 'none' }" :data-element-id="element.id" :data-autofit="element.textAutoFit === 'shape' ? 'shape' : element.textAutoFitDynamic ? 'normal' : undefined" :data-text-overflow="supportsTextVerticalEllipsis(element) ? 'ellipsis' : undefined" :style="textFrameStyle(element)">
        <div class="text-bounds" :style="textBoundsStyle(element)">
          <div
            v-for="(entry, index) in paragraphEntries(element.paragraphs)"
            :key="index"
            class="text-paragraph"
            :style="[paragraphStyle(entry.paragraph, effectiveTextFontScale(element), element.textLineSpacingReduction, usesRightToLeftColumns(element) ? 'ltr' : undefined), paragraphAnimationStyle(element, index)]"
          >
            <img v-if="entry.paragraph.bullet?.kind === 'image' && entry.paragraph.bullet.url" class="paragraph-picture-bullet" :src="entry.paragraph.bullet.url" :style="paragraphBulletStyle(entry.paragraph)" alt="" aria-hidden="true">
            <span v-else-if="entry.marker" class="paragraph-marker" :style="paragraphBulletStyle(entry.paragraph)" aria-hidden="true">{{ entry.marker }}</span>
            <template v-for="(run, runIndex) in entry.paragraph.runs" :key="runIndex">
              <component
                v-for="(segment, segmentIndex) in runSegments(run, element, index, runIndex)"
                :key="segmentIndex"
                :is="canActivateHyperlink(run.hyperlink) ? 'a' : 'span'"
                :href="canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : undefined"
                :title="run.hyperlink?.tooltip"
                :data-pptx-hyperlink="canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : undefined"
                :class="['text-run', { 'text-run-hyperlink': canActivateHyperlink(run.hyperlink), 'text-run-break': segment.text === '\n', 'text-run-fill': !segment.separateDecorations && runHasTextFill(element, run, index, segment.start, segment.end), 'text-run-highlight': Boolean(run.highlightColor) }]"
                :data-char-start="segment.start"
                :data-char-end="segment.end"
                :style="[runStyle(element, run, segment.script, effectiveTextFontScale(element), index, segment.start, segment.end, segment.separateDecorations, 'text-' + runIndex), textColorAnimationStyle(element, index, segment.start, segment.end, run.color), textRangeOpacityAnimationStyle(element, index, segment.start, segment.end), characterAnimationStyle(element, index, segment.start, segment.end), textRunAnimationStyle(element, index, segment.start, segment.end, run, effectiveTextFontScale(element), segment.script, segment.separateDecorations, 'text-' + runIndex), textTransformAnimationStyle(element, index, segment.start, segment.end)]"
              >
                <svg v-if="textInnerShadows(run).length" class="text-run-filter-defs" width="0" height="0" aria-hidden="true" focusable="false">
                  <defs>
                    <filter :id="textInnerShadowFilterId(element, index, segment.start, segment.end, 'text-' + runIndex)" x="-10%" y="-10%" width="120%" height="120%" filterUnits="objectBoundingBox" primitiveUnits="userSpaceOnUse" color-interpolation-filters="sRGB">
                      <template v-for="(shadow, shadowIndex) in textInnerShadows(run)" :key="shadowIndex">
                        <feGaussianBlur in="SourceAlpha" :stdDeviation="Math.max(0, shadow.blur * scale / 2)" :result="'inner-blur-' + shadowIndex" />
                        <feOffset :in="'inner-blur-' + shadowIndex" :dx="textInnerShadowOffset(element, shadow, index, segment.start, segment.end, shadowIndex).dx" :dy="textInnerShadowOffset(element, shadow, index, segment.start, segment.end, shadowIndex).dy" :result="'inner-offset-' + shadowIndex" />
                        <feComposite in="SourceAlpha" :in2="'inner-offset-' + shadowIndex" operator="out" :result="'inner-mask-' + shadowIndex" />
                        <feFlood :flood-color="textInnerShadowColor(element, shadow, index, segment.start, segment.end)" :result="'inner-color-' + shadowIndex" />
                        <feComposite :in="'inner-color-' + shadowIndex" :in2="'inner-mask-' + shadowIndex" operator="in" :result="'inner-shadow-' + shadowIndex" />
                      </template>
                      <feMerge>
                        <feMergeNode v-for="(_, shadowIndex) in textInnerShadows(run)" :key="shadowIndex" :in="'inner-shadow-' + shadowIndex" />
                        <feMergeNode in="SourceGraphic" />
                      </feMerge>
                    </filter>
                  </defs>
                </svg>
                  <span v-if="segment.separateDecorations" class="text-run-underline" :class="{ 'text-run-fill': runHasTextFill(element, run, index, segment.start, segment.end), 'text-run-highlight': Boolean(run.highlightColor) }" :style="textUnderlineStyle(run)">{{ segment.text }}</span>
                <template v-else>{{ segment.text }}</template>
              </component>
            </template>
          </div>
        </div>
      </div>
      <svg v-if="element.kind === 'line'" class="line-render" :style="lineRenderStyle(element)" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true">
        <defs>
          <pattern v-if="element.strokePattern" :id="customGeometryStrokePatternId(element)" patternUnits="userSpaceOnUse" width="1000" height="1000"><foreignObject width="1000" height="1000"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%" :style="{ background: element.strokePattern }" /></foreignObject></pattern>
          <marker v-if="animationLineEnd(element, 'head').type !== 'none'" :id="lineEndId(element, 'head')" :markerWidth="lineEndMarkerSize(element, 'head').width" :markerHeight="lineEndMarkerSize(element, 'head').height" viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto-start-reverse">
            <path :d="lineEndPath(animationLineEnd(element, 'head').type)" :fill="lineEndUsesStroke(animationLineEnd(element, 'head').type) ? 'none' : lineStrokeColor(element)" :stroke="lineEndUsesStroke(animationLineEnd(element, 'head').type) ? lineStrokeColor(element) : 'none'" />
          </marker>
          <marker v-if="animationLineEnd(element, 'tail').type !== 'none'" :id="lineEndId(element, 'tail')" :markerWidth="lineEndMarkerSize(element, 'tail').width" :markerHeight="lineEndMarkerSize(element, 'tail').height" viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto">
            <path :d="lineEndPath(animationLineEnd(element, 'tail').type)" :fill="lineEndUsesStroke(animationLineEnd(element, 'tail').type) ? 'none' : lineStrokeColor(element)" :stroke="lineEndUsesStroke(animationLineEnd(element, 'tail').type) ? lineStrokeColor(element) : 'none'" />
          </marker>
        </defs>
        <path v-if="element.linePath" :d="element.linePath" fill="none" :stroke="lineStrokeColor(element)" :stroke-width="customGeometryStrokeWidth(element)" :style="customGeometryStrokeStyle(element)" :marker-start="lineEndUrl(element, 'head')" :marker-end="lineEndUrl(element, 'tail')" vector-effect="non-scaling-stroke" />
        <line v-else x1="0" :y1="element.geometry === 'lineInv' ? 1000 : 0" :x2="element.width > 0 ? 1000 : 0" :y2="element.height > 0 ? (element.geometry === 'lineInv' ? 0 : 1000) : 0" :stroke="lineStrokeColor(element)" :stroke-width="customGeometryStrokeWidth(element)" :style="customGeometryStrokeStyle(element)" :marker-start="lineEndUrl(element, 'head')" :marker-end="lineEndUrl(element, 'tail')" vector-effect="non-scaling-stroke" />
      </svg>
    </div>
    </div>
    <audio
      v-for="action in embeddedSoundActions"
      :key="action.mediaActionKey || action.targetId"
      class="animation-sound"
      data-animation-sound
      :data-media-id="action.targetId"
      :src="action.mediaSourceUrl"
      preload="none"
      aria-hidden="true"
      @ended="finishTimedMedia(action.targetId, true, true)"
    />
    <span v-if="backgroundTileWarning" class="media-warning" role="status">{{ t(backgroundTileWarning) }}</span>
    <span v-if="textOverflowWarning && (!thumbnail || playbackPreview)" class="media-warning" role="status">{{ t(textOverflowWarning) }}</span>
  </div>
</template>

<style scoped>
.image-viewport { position: relative; width: 100%; height: 100%; overflow: hidden; }
.image-tone-defs { position: absolute; width: 0; height: 0; overflow: visible; pointer-events: none; }
.picture-inner-shadow { position: absolute; inset: 0; z-index: 1; pointer-events: none; }
.custom-geometry, .preset-outline { position: absolute; inset: 0; width: 100%; height: 100%; overflow: visible; pointer-events: none; }
.slide-image { position: absolute; max-width: none; }
.slide-video { position: absolute; max-width: none; object-fit: fill; }
.animation-sound { display: none; }
.audio-viewport { position: relative; width: 100%; height: 100%; display: grid; align-items: center; }
.slide-audio { width: 100%; max-width: 100%; }
.media-warning { position: absolute; inset: auto 0 0; padding: 4px 8px; background: #7f1d1d; color: white; font: 12px/1.3 sans-serif; }
.is-animation-trigger { cursor: pointer; }
.line-render { position: absolute; left: 0; top: 0; overflow: visible; pointer-events: none; }
.slide-chart { display: block; width: 100%; height: 100%; overflow: hidden; font: 20px/1.2 Arial, sans-serif; }
.slide-chart .chart-title { fill: #1f2937; font-size: 26px; font-weight: 600; }
.slide-chart .chart-label { fill: #475569; font-size: 18px; }
.slide-chart .chart-data-label { fill: #1f2937; font-size: 18px; font-weight: 600; }
.slide-table { width: 100%; height: 100%; border-collapse: collapse; table-layout: fixed; }
.slide-table td { box-sizing: border-box; overflow: hidden; }
.text-run-break { white-space: pre; }
.text-run-filter-defs { position: absolute; width: 0; height: 0; overflow: visible; pointer-events: none; }
.text-run-fill { color: var(--pptx-text-fill-color, currentColor); }
.text-run-highlight { background-color: var(--pptx-text-highlight); -webkit-box-decoration-break: clone; box-decoration-break: clone; }
@supports (background-clip: text) or (-webkit-background-clip: text) {
  .text-run-fill { background: var(--pptx-text-fill); background-clip: text; -webkit-background-clip: text; color: transparent !important; -webkit-text-fill-color: transparent; }
  .text-run-fill.text-run-highlight { background: var(--pptx-text-highlight-background); background-clip: var(--pptx-text-fill-clip); -webkit-background-clip: var(--pptx-text-fill-clip); }
}
.paragraph-marker { display: inline-block; width: 1em; margin-left: -1em; margin-right: .45em; text-align: right; white-space: pre; }
.paragraph-picture-bullet { display: inline-block; width: 1em; height: 1em; margin-left: -1em; margin-right: .45em; object-fit: contain; vertical-align: middle; }
.is-thumbnail .text-frame { overflow: hidden; }
@media (prefers-reduced-motion: reduce) { .slide-element { transition: none !important; } }
</style>
