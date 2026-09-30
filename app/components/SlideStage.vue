<script setup lang="ts">
import { computed, nextTick, reactive, useId, watch, type CSSProperties } from 'vue'
import { formatChartNumber } from '~/utils/chart-number-format'
import { hslToRgb } from '~/utils/color'
import { getPptxChartStacking } from '~/utils/pptx'
import type { PptxAnimation, PptxChart, PptxChartDataLabelPosition, PptxChartTickLabelPosition, PptxChartTickMark, PptxElement, PptxHyperlink, PptxParagraph, PptxSlide, PptxTableCell, PptxTriggerPlayback } from '~/utils/pptx'
import { inverseTimingProgress, timingProgress } from '~/utils/animation-timing'
import { barnClipPath, clipPathMaskImage, dissolveClipPath, gridTransitionClipPath, gridTransitionMaskImage, randomBarClipPath, shapeEffectClipPath, stripsClipPath, wheelClipPath } from '~/utils/transitionMasks'

const EMU_PER_POINT = 12_700
const EAST_ASIAN_SCRIPT = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}]/u
const COMPLEX_SCRIPT = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Thai}\p{Script=Lao}\p{Script=Tibetan}\p{Script=Myanmar}\p{Script=Khmer}\p{Script=Mongolian}]/u
const customGradientScope = `pptx-${useId()}`.replace(/[^a-zA-Z0-9_-]/g, '_')

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
interface CustomSvgRadialGradient {
  id: string
  kind: 'radial'
  cx: number
  cy: number
  r: number
  stops: CustomSvgGradientStop[]
}
type CustomSvgGradient = CustomSvgLinearGradient | CustomSvgRadialGradient

interface ChartVisual {
  plot: { left: number; top: number; width: number; height: number }
  ticks: Array<{ value: number; label: string; showLabel: boolean; x: number; y: number; labelX: number; labelY: number; labelAnchor: string }>
  axisTicks: Array<{ axis: 'category' | 'value'; x1: number; y1: number; x2: number; y2: number }>
  categories: Array<{ text: string; x: number; y: number; anchor: string }>
  bars: Array<{ x: number; y: number; width: number; height: number; color: string }>
  lines: Array<{ d: string; color: string }>
  markers: Array<{ x: number; y: number; color: string }>
  slices: Array<{ d: string; color: string }>
  legend: Array<{ text: string; x: number; y: number; color: string }>
  legendClip?: { id: string; x: number; y: number; width: number; height: number }
  dataLabels: Array<{ text: string; lines: string[]; x: number; y: number; anchor: string }>
  baseline: { x1: number; y1: number; x2: number; y2: number }
  valueAxisLine?: { x1: number; y1: number; x2: number; y2: number }
}

type ChartLegendPosition = NonNullable<PptxChart['legend']>['position']

function defaultLegendOrigin(position: ChartLegendPosition | undefined, axis: 'x' | 'y') {
  if (axis === 'x') return position === 'left' ? 88 : position === 'right' ? 780 : position === 'topRight' ? 760 : 88
  return position === 'left' || position === 'right' ? 105 : position === 'top' ? 77 : position === 'topRight' ? 92 : 530
}

function makeChartVisual(chart: PptxChart, elementId: string): ChartVisual {
  const position = chart.legend?.position
  const plotPosition = chart.legend?.overlay ? undefined : position
  const plot = plotPosition === 'left' ? { left: 330, top: 75, width: 608, height: 390 }
    : plotPosition === 'right' ? { left: 88, top: 75, width: 650, height: 390 }
      : plotPosition === 'top' ? { left: 88, top: 145, width: 850, height: 320 }
        : { left: 88, top: 75, width: 850, height: 390 }
  const values = chart.series.flatMap(series => series.values.filter((value): value is number => value !== null && value !== undefined))
  const dataMin = Math.min(0, ...values)
  const dataMax = Math.max(0, ...values)
  const minimum = chart.valueAxis?.min ?? dataMin
  const upper = chart.valueAxis?.max ?? (dataMax === dataMin ? dataMin + 1 : dataMax)
  const categoryCount = Math.max(chart.categories.length, 1)
  const horizontalValueAxis = chart.direction === 'horizontal'
  const valueAxisSide = chart.valueAxis?.position ?? (horizontalValueAxis ? 'b' : 'l')
  const categoryAxisSide = chart.categoryAxisPosition ?? (horizontalValueAxis ? 'l' : 'b')
  const valueTickLabelPosition = chart.valueAxis?.tickLabelPosition ?? 'nextTo'
  const categoryTickLabelPosition = chart.categoryAxisTickLabelPosition ?? 'nextTo'
  const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value))
  const axisFraction = (value: number) => {
    const fraction = (value - minimum) / (upper - minimum)
    return chart.valueAxis?.reverse ? 1 - fraction : fraction
  }
  const valuePosition = (value: number) => plot.top + plot.height * (1 - axisFraction(value))
  const categoryCrossingFraction = chart.categoryAxisCrossesAt !== undefined
    ? (clamp(chart.categoryAxisCrossesAt, 1, categoryCount) - 0.5) / categoryCount
    : chart.valueAxis?.crosses === 'min' ? 0 : chart.valueAxis?.crosses === 'max' ? 1 : undefined
  const valueAxisCrossingPosition = categoryCrossingFraction === undefined ? undefined
    : horizontalValueAxis ? plot.top + plot.height * categoryCrossingFraction
      : plot.left + plot.width * categoryCrossingFraction
  const categoryAxisCrossingValue = chart.valueAxis?.crossesAt ?? (chart.categoryAxisCrosses === 'min' ? minimum : chart.categoryAxisCrosses === 'max' ? upper : clamp(0, minimum, upper))
  const categoryAxisCrossingCoordinate = horizontalValueAxis
    ? plot.left + plot.width * axisFraction(clamp(categoryAxisCrossingValue, minimum, upper))
    : valuePosition(clamp(categoryAxisCrossingValue, minimum, upper))
  const valueAxisLineCoordinate = valueAxisCrossingPosition ?? (horizontalValueAxis
    ? valueAxisSide === 't' ? plot.top : plot.top + plot.height
    : valueAxisSide === 'r' ? plot.left + plot.width : plot.left)
  const valueAxisLine = horizontalValueAxis
    ? { x1: plot.left, y1: valueAxisLineCoordinate, x2: plot.left + plot.width, y2: valueAxisLineCoordinate }
    : { x1: valueAxisLineCoordinate, y1: plot.top, x2: valueAxisLineCoordinate, y2: plot.top + plot.height }
  const unit = chart.valueAxis?.majorUnit
  const tickValues = unit
    ? Array.from({ length: Math.min(101, Math.floor((upper - minimum) / unit + 1e-9) + 1) }, (_, index) => Number((minimum + index * unit).toPrecision(12)))
    : Array.from({ length: 5 }, (_, index) => minimum + (upper - minimum) * index / 4)
  const ticks = tickValues.map(value => {
    const x = plot.left + plot.width * axisFraction(value)
    const y = valuePosition(value)
    return {
      value,
      label: formatChartNumber(value, chart.valueAxis?.numberFormat),
      showLabel: valueTickLabelPosition !== 'none' && !chart.valueAxis?.deleted,
      x,
      y,
      labelX: horizontalValueAxis ? x : valueTickLabelPosition === 'high' ? plot.left + plot.width + 10 : valueTickLabelPosition === 'low' ? plot.left - 10 : valueAxisLineCoordinate + (valueAxisSide === 'r' ? 10 : -10),
      labelY: horizontalValueAxis ? valueTickLabelPosition === 'high' ? plot.top + plot.height + 22 : valueTickLabelPosition === 'low' ? plot.top - 10 : valueAxisLineCoordinate + (valueAxisSide === 't' ? -10 : 22) : y + 5,
      labelAnchor: horizontalValueAxis ? 'middle' : valueTickLabelPosition === 'high' ? 'start' : valueTickLabelPosition === 'low' ? 'end' : valueAxisSide === 'r' ? 'start' : 'end',
    }
  })
  const axisTicks: ChartVisual['axisTicks'] = []
  const addAxisTick = (axis: 'category' | 'value', mark: PptxChartTickMark | undefined, x: number, y: number, horizontal: boolean, inwardDirection: number) => {
    if (!mark || mark === 'none') return
    const start = mark === 'cross' ? -5 : 0
    const end = mark === 'cross' ? 5 : mark === 'in' ? 5 : -5
    axisTicks.push({
      axis,
      x1: horizontal ? x + start * inwardDirection : x,
      y1: horizontal ? y : y + start * inwardDirection,
      x2: horizontal ? x + end * inwardDirection : x,
      y2: horizontal ? y : y + end * inwardDirection,
    })
  }
  const valueTickInwardDirection = horizontalValueAxis ? valueAxisSide === 't' ? 1 : -1 : valueAxisSide === 'l' ? 1 : -1
  if (!chart.valueAxis?.deleted) for (const tick of ticks) addAxisTick('value', chart.valueAxis?.majorTickMark, horizontalValueAxis ? tick.x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : tick.y, !horizontalValueAxis, valueTickInwardDirection)
  const categoryTickInwardDirection = horizontalValueAxis ? categoryAxisSide === 'l' ? 1 : -1 : categoryAxisSide === 't' ? 1 : -1
  if (!chart.categoryAxisDeleted && chart.categoryAxisMajorTickMark && chart.categoryAxisMajorTickMark !== 'none') {
    for (let index = 0; index < categoryCount; index++) {
      const centerX = plot.left + plot.width * (index + 0.5) / categoryCount
      const centerY = plot.top + plot.height * (index + 0.5) / categoryCount
      addAxisTick('category', chart.categoryAxisMajorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : centerX, horizontalValueAxis ? centerY : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection)
    }
  }
  const categories: ChartVisual['categories'] = []
  const bars: ChartVisual['bars'] = []
  const lines: ChartVisual['lines'] = []
  const markers: ChartVisual['markers'] = []
  const slices: ChartVisual['slices'] = []
  const legend: ChartVisual['legend'] = []
  const dataLabels: ChartVisual['dataLabels'] = []
  const formatValue = (value: number, series: PptxChart['series'][number], index: number) => formatChartNumber(value, series.valueNumberFormatOverrides[index] ?? series.numberFormat ?? chart.numberFormat)
  const labelPosition = (series: PptxChart['series'][number], index: number, fallback: PptxChartDataLabelPosition) => series.dataLabelPositionOverrides[index] ?? series.dataLabelPosition ?? chart.dataLabelPosition ?? fallback
  const pushDataLabel = (text: string, x: number, y: number, anchor: string) => dataLabels.push({ text, lines: text.split(/\r?\n/), x, y, anchor })
  const formatDataLabel = (series: PptxChart['series'][number], index: number, value: number, category: string, total?: number) => {
    const showValue = series.valueLabelOverrides[index] ?? series.showValueLabels
    const showCategory = series.categoryNameLabelOverrides[index] ?? series.showCategoryNameLabels
    const showSeries = series.seriesNameLabelOverrides[index] ?? series.showSeriesNameLabels
    const isPieLike = chart.type === 'pie' || chart.type === 'doughnut'
    const showPercent = isPieLike && (series.percentLabelOverrides[index] ?? series.showPercentLabels)
    const parts = [
      showValue ? formatValue(value, series, index) : undefined,
      showCategory ? category : undefined,
      showSeries ? series.name : undefined,
      showPercent && total ? formatChartNumber(value / total, series.valueNumberFormatOverrides[index] ?? series.numberFormat ?? chart.numberFormat ?? '0.0%') : undefined,
    ].filter((part): part is string => part !== undefined)
    if (!parts.length) return undefined
    const defaultSeparator = isPieLike && !showValue && showCategory && !showSeries && showPercent ? '\n' : ','
    return parts.join(series.labelSeparatorOverrides[index] ?? series.labelSeparator ?? chart.labelSeparator ?? defaultSeparator)
  }

  if (chart.type === 'pie' || chart.type === 'doughnut') {
    const series = chart.series[0]!
    const centerX = plot.left + plot.width / 2
    const centerY = plot.top + plot.height / 2
    const radius = Math.min(plot.width, plot.height) * 0.47
    const innerRadius = chart.type === 'doughnut' ? radius * (chart.holeSize ?? 50) / 100 : 0
    const points = chart.categories.flatMap((text, index) => {
      const value = series.values[index]
      return value != null && value > 0 ? [{ text, index, value, color: series.pointColors[index] || chart.palette[index % chart.palette.length] || series.color }] : []
    })
    const total = points.reduce((sum, point) => sum + point.value, 0)
    let angle = -Math.PI / 2 + (chart.firstSliceAngle ?? 0) * Math.PI / 180
    for (const point of points) {
      const sweep = point.value / total * Math.PI * 2
      const startX = centerX + Math.cos(angle) * radius
      const startY = centerY + Math.sin(angle) * radius
      const endAngle = angle + sweep
      const endX = centerX + Math.cos(endAngle) * radius
      const endY = centerY + Math.sin(endAngle) * radius
      const endInnerX = centerX + Math.cos(endAngle) * innerRadius
      const endInnerY = centerY + Math.sin(endAngle) * innerRadius
      const startInnerX = centerX + Math.cos(angle) * innerRadius
      const startInnerY = centerY + Math.sin(angle) * innerRadius
      const fullCircle = sweep >= Math.PI * 2 - 0.0001
      const d = innerRadius > 0
        ? fullCircle
          ? `M ${startX} ${startY} A ${radius} ${radius} 0 1 1 ${centerX - Math.cos(angle) * radius} ${centerY - Math.sin(angle) * radius} A ${radius} ${radius} 0 1 1 ${startX} ${startY} L ${startInnerX} ${startInnerY} A ${innerRadius} ${innerRadius} 0 1 0 ${centerX - Math.cos(angle) * innerRadius} ${centerY - Math.sin(angle) * innerRadius} A ${innerRadius} ${innerRadius} 0 1 0 ${startInnerX} ${startInnerY} Z`
          : `M ${startX} ${startY} A ${radius} ${radius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${endX} ${endY} L ${endInnerX} ${endInnerY} A ${innerRadius} ${innerRadius} 0 ${sweep > Math.PI ? 1 : 0} 0 ${startInnerX} ${startInnerY} Z`
        : fullCircle
          ? `M ${centerX} ${centerY} L ${startX} ${startY} A ${radius} ${radius} 0 1 1 ${centerX - Math.cos(angle) * radius} ${centerY - Math.sin(angle) * radius} A ${radius} ${radius} 0 1 1 ${startX} ${startY} Z`
          : `M ${centerX} ${centerY} L ${startX} ${startY} A ${radius} ${radius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${endX} ${endY} Z`
      slices.push({ d, color: point.color })
      const label = formatDataLabel(series, point.index, point.value, point.text, total)
      if (label !== undefined) {
        const labelAngle = angle + sweep / 2
        const position = labelPosition(series, point.index, 'ctr')
        const labelRadius = chart.type === 'doughnut'
          ? position === 'outEnd' ? radius * 1.12
            : position === 'inEnd' ? innerRadius + (radius - innerRadius) * 0.82
              : position === 'inBase' ? innerRadius + (radius - innerRadius) * 0.18
                : (radius + innerRadius) / 2
          : radius * (position === 'outEnd' ? 1.12 : position === 'inEnd' ? 0.82 : position === 'inBase' ? 0.24 : position === 'bestFit' ? 0.62 : 0.68)
        const x = position === 'l' ? centerX - radius * 0.75 : position === 'r' ? centerX + radius * 0.75 : centerX + Math.cos(labelAngle) * labelRadius
        const y = position === 't' ? centerY - radius * 0.78 : position === 'b' ? centerY + radius * 0.78 : centerY + Math.sin(labelAngle) * labelRadius + 6
        const anchor = position === 'l' ? 'end' : position === 'r' ? 'start' : 'middle'
        pushDataLabel(label, x, y, anchor)
      }
      angle = endAngle
    }
    points.slice(0, 12).forEach(point => {
      if (chart.legend?.hiddenEntries?.includes(point.index)) return
      legend.push({ text: point.text, x: 590, y: 105 + legend.length * 28, color: point.color })
    })
  } else {
    const seriesCount = Math.max(chart.series.length, 1)
    for (let index = 0; index < categoryCount; index++) {
      const centerX = plot.left + plot.width * (index + 0.5) / categoryCount
      const centerY = plot.top + plot.height * (index + 0.5) / categoryCount
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
        categories.push({ text: chart.categories[index] || '', ...categoryLabel })
      }
      chart.series.forEach((series, seriesIndex) => {
        const value = series.values[index]
        if (value === null || value === undefined) return
        const color = series.color
        if (chart.type === 'line') {
          const x = centerX
          const y = valuePosition(value)
          const previousValue = index > 0 ? series.values[index - 1] : null
          const previous = previousValue == null ? undefined : ({ x: plot.left + plot.width * (index - 0.5) / categoryCount, y: valuePosition(previousValue) })
          lines.push({ d: previous ? `M ${previous.x} ${previous.y} L ${x} ${y}` : `M ${x} ${y}`, color })
          markers.push({ x, y, color })
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, 't')
            const labelX = position === 'l' ? x - 10 : position === 'r' ? x + 10 : x
            const labelY = position === 'b' ? y + 24 : position === 'ctr' || position === 'inBase' || position === 'inEnd' ? y + 7 : y - 12
            const anchor = position === 'l' ? 'end' : position === 'r' ? 'start' : 'middle'
            pushDataLabel(label, labelX, labelY, anchor)
          }
        } else if (chart.direction === 'horizontal') {
          const groupHeight = plot.height / categoryCount * 0.72
          const barHeight = groupHeight / seriesCount
          const barY = centerY - groupHeight / 2 + seriesIndex * barHeight
          const zeroX = plot.left + plot.width * axisFraction(Math.max(minimum, Math.min(upper, 0)))
          const valueX = plot.left + plot.width * axisFraction(value)
          bars.push({ x: Math.min(zeroX, valueX), y: barY, width: Math.max(1, Math.abs(valueX - zeroX)), height: barHeight * 0.86, color })
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, 'outEnd')
            const direction = value >= 0 ? 1 : -1
            const left = Math.min(zeroX, valueX)
            const right = Math.max(zeroX, valueX)
            let labelX = valueX + direction * 8
            let labelY = barY + barHeight * 0.68
            let anchor = direction > 0 ? 'start' : 'end'
            if (position === 'ctr') { labelX = (zeroX + valueX) / 2; anchor = 'middle' }
            else if (position === 'inBase') { labelX = zeroX + direction * 20; anchor = direction > 0 ? 'start' : 'end' }
            else if (position === 'inEnd') { labelX = valueX - direction * 20; anchor = direction > 0 ? 'end' : 'start' }
            else if (position === 'l') { labelX = left - 8; anchor = 'end' }
            else if (position === 'r') { labelX = right + 8; anchor = 'start' }
            else if (position === 't') { labelX = (zeroX + valueX) / 2; labelY = barY - 8; anchor = 'middle' }
            else if (position === 'b') { labelX = (zeroX + valueX) / 2; labelY = barY + barHeight + 15; anchor = 'middle' }
            pushDataLabel(label, labelX, labelY, anchor)
          }
        } else {
          const groupWidth = plot.width / categoryCount * 0.72
          const barWidth = groupWidth / seriesCount
          const zeroY = valuePosition(Math.max(minimum, Math.min(upper, 0)))
          const valueY = valuePosition(value)
          const barX = centerX - groupWidth / 2 + seriesIndex * barWidth
          bars.push({ x: barX, y: Math.min(zeroY, valueY), width: barWidth * 0.86, height: Math.max(1, Math.abs(valueY - zeroY)), color })
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1))
          if (label !== undefined) {
            const position = labelPosition(series, index, 'outEnd')
            const top = Math.min(zeroY, valueY)
            const bottom = Math.max(zeroY, valueY)
            let labelX = barX + barWidth * 0.43
            let labelY = value >= 0 ? top - 8 : bottom + 20
            let anchor = 'middle'
            if (position === 'ctr') labelY = (zeroY + valueY) / 2 + 6
            else if (position === 'inBase') labelY = value >= 0 ? zeroY - 8 : zeroY + 18
            else if (position === 'inEnd') labelY = value >= 0 ? valueY + 18 : valueY - 4
            else if (position === 'b') labelY = bottom + 20
            else if (position === 'l') { labelX = barX - 8; labelY = (zeroY + valueY) / 2 + 6; anchor = 'end' }
            else if (position === 'r') { labelX = barX + barWidth * 0.86 + 8; labelY = (zeroY + valueY) / 2 + 6; anchor = 'start' }
            pushDataLabel(label, labelX, labelY, anchor)
          }
        }
      })
    }
    chart.series.slice(0, 12).forEach(series => {
      if (chart.legend?.hiddenEntries?.includes(series.legendIndex)) return
      legend.push({ text: series.name, x: 88 + (legend.length % 4) * 220, y: 530 + Math.floor(legend.length / 4) * 22, color: series.color })
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
  const baseline = horizontalValueAxis
    ? { x1: categoryAxisCrossingCoordinate, y1: plot.top, x2: categoryAxisCrossingCoordinate, y2: plot.top + plot.height }
    : { x1: plot.left, y1: categoryAxisCrossingCoordinate, x2: plot.left + plot.width, y2: categoryAxisCrossingCoordinate }
  return { plot, ticks, axisTicks, categories, bars, lines, markers, slices, legend, ...(legendClip ? { legendClip } : {}), dataLabels, baseline, valueAxisLine }
}

const props = withDefaults(defineProps<{
  slide: PptxSlide
  width: number
  height: number
  number: number
  animationStep?: number
  thumbnail?: boolean
  playbackPreview?: boolean
  triggerPlayback?: PptxTriggerPlayback[]
}>(), { animationStep: 0, thumbnail: false, playbackPreview: false })
const emit = defineEmits<{
  'trigger-playback': [playback: PptxTriggerPlayback[]]
  'activate-hyperlink': [hyperlink: PptxHyperlink]
}>()

const backgroundImageSize = ref<{ sourceUrl: string; url: string; width: number; height: number; repeatX: number; repeatY: number }>()
const backgroundTileWarning = ref('')
watch(() => [props.slide.backgroundImage?.url, Boolean(props.slide.backgroundImage?.tile), props.slide.backgroundImage?.tile?.flip || 'none'] as const, ([url, tiled, flip], _oldValue, onCleanup) => {
  backgroundImageSize.value = undefined
  backgroundTileWarning.value = ''
  if (!import.meta.client || !url || !tiled) return
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
    if (current) backgroundTileWarning.value = 'The tiled slide background image could not be decoded.'
  }
  probe.onload = () => {
    if (!current || probe.naturalWidth < 1 || probe.naturalHeight < 1) return
    const width = probe.naturalWidth
    const height = probe.naturalHeight
    const flipX = flip === 'x' || flip === 'xy'
    const flipY = flip === 'y' || flip === 'xy'
    const imageSize = { sourceUrl: url, url, width, height, repeatX: flipX ? 2 : 1, repeatY: flipY ? 2 : 1 }
    if (!flipX && !flipY) {
      backgroundImageSize.value = imageSize
      return
    }
    // ponytail: cap synthesized flip patterns at 16 MP; downsample if real decks hit this ceiling.
    if (width * imageSize.repeatX > 32_767 || height * imageSize.repeatY > 32_767 || width * height * imageSize.repeatX * imageSize.repeatY > 16_777_216) {
      backgroundTileWarning.value = 'The tiled background flip exceeds the local canvas limit; the unflipped image is used.'
      backgroundImageSize.value = { ...imageSize, repeatX: 1, repeatY: 1 }
      return
    }
    const canvas = document.createElement('canvas')
    canvas.width = width * imageSize.repeatX
    canvas.height = height * imageSize.repeatY
    const context = canvas.getContext('2d')
    if (!context) {
      backgroundTileWarning.value = 'The browser could not render the tiled background flip; the unflipped image is used.'
      backgroundImageSize.value = { ...imageSize, repeatX: 1, repeatY: 1 }
      return
    }
    const drawTile = (x: number, y: number, mirrorX: boolean, mirrorY: boolean) => {
      context.save()
      context.translate(x + (mirrorX ? width : 0), y + (mirrorY ? height : 0))
      context.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1)
      context.drawImage(probe, 0, 0, width, height)
      context.restore()
    }
    drawTile(0, 0, false, false)
    if (flipX) drawTile(width, 0, true, false)
    if (flipY) drawTile(0, height, false, true)
    if (flipX && flipY) drawTile(width, height, true, true)
    canvas.toBlob(blob => {
      if (!current) return
      if (!blob) {
        backgroundTileWarning.value = 'The browser could not encode the tiled background flip; the unflipped image is used.'
        backgroundImageSize.value = { ...imageSize, repeatX: 1, repeatY: 1 }
        return
      }
      generatedUrl = URL.createObjectURL(blob)
      backgroundImageSize.value = { ...imageSize, url: generatedUrl }
    }, 'image/png')
  }
  probe.src = url
}, { immediate: true })

const frameStyle = computed<CSSProperties>(() => {
  const image = props.slide.backgroundImage
  if (!image) return { aspectRatio: `${props.width} / ${props.height}`, background: props.slide.background }
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
    aspectRatio: `${props.width} / ${props.height}`,
    background: props.slide.background,
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
const animationClock = ref(0)
const reduceMotion = ref(false)
const invalidMediaTrimIds = ref<string[]>([])
const invalidMediaFadeIds = ref<string[]>([])
const blockedMediaPlaybackIds = ref<string[]>([])
const mediaFadeFrames = new Map<HTMLMediaElement, number>()
const mediaFadeStates = new WeakMap<HTMLMediaElement, { baseVolume: number; appliedVolume: number; appliedFactor: number; baseMuted: boolean; appliedMuted: boolean }>()
const executedMediaActions = new Set<number>()
const completedMediaActions = reactive(new Map<string, number>())
const mediaActionEndTimers = new Map<string, ReturnType<typeof setTimeout>>()
let automaticStartedAt: number | undefined
let stepStartedAt = new Map<number, number>()
let triggerSteps = new Map<number, number>()
let triggerStartedAt = new Map<string, number>()
let triggerPlaybackByKey = new Map<string, PptxTriggerPlayback>()
let observer: ResizeObserver | undefined
let autoFitFrame = 0
let animationFrame = 0
let animationTimer: ReturnType<typeof setTimeout> | undefined
let motionPreference: MediaQueryList | undefined
const scale = computed(() => stageWidth.value / Math.max(props.width, 1))
const chartVisuals = computed(() => new Map(props.slide.elements.flatMap(element => element.chart ? [[element.id, makeChartVisual(element.chart, element.id)] as const] : [])))

function scheduleShapeAutoFit() {
  if (!stage.value || autoFitFrame || scale.value <= 0) return
  autoFitFrame = requestAnimationFrame(() => {
    autoFitFrame = 0
    if (!stage.value) return
    const elements = new Map(props.slide.elements.map(element => [element.id, element]))
    let changed = false
    for (const frame of stage.value.querySelectorAll<HTMLElement>('.text-frame[data-autofit="shape"]')) {
      const id = frame.dataset.elementId
      const element = id ? elements.get(id) : undefined
      if (!id || !element) continue
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
    }
    for (const frame of stage.value.querySelectorAll<HTMLElement>('.text-frame[data-autofit="normal"]')) {
      const id = frame.dataset.elementId
      const element = id ? elements.get(id) : undefined
      if (!id || !element || !frame.clientWidth || !frame.clientHeight) continue
      const overflowWidth = frame.scrollWidth > frame.clientWidth + 1 ? frame.clientWidth / frame.scrollWidth : 1
      const overflowHeight = frame.scrollHeight > frame.clientHeight + 1 ? frame.clientHeight / frame.scrollHeight : 1
      const key = `${props.slide.id}:${id}`
      const current = normalAutoFitScales.get(key) ?? element.textFontScale ?? 1
      const fitted = Math.max(0.05, current * Math.min(overflowWidth, overflowHeight) * 0.98)
      if (fitted < current - 0.001) {
        normalAutoFitScales.set(key, fitted)
        changed = true
      }
    }
    for (const cell of stage.value.querySelectorAll<HTMLTableCellElement>('.slide-table td[data-autofit="normal"]')) {
      const elementNode = cell.closest<HTMLElement>('.slide-element')
      const id = elementNode?.dataset.elementId
      const rowIndex = Number(cell.dataset.rowIndex)
      const cellIndex = Number(cell.dataset.cellIndex)
      const element = id ? elements.get(id) : undefined
      const model = element?.table?.rows[rowIndex]?.cells[cellIndex]
      if (!id || !model || !cell.clientWidth || !cell.clientHeight) continue
      const overflowWidth = cell.scrollWidth > cell.clientWidth + 1 ? cell.clientWidth / cell.scrollWidth : 1
      const overflowHeight = cell.scrollHeight > cell.clientHeight + 1 ? cell.clientHeight / cell.scrollHeight : 1
      const key = tableCellAutoFitKey(id, rowIndex, cellIndex)
      const current = normalAutoFitScales.get(key) ?? model.textFontScale ?? 1
      const fitted = Math.max(0.05, current * Math.min(overflowWidth, overflowHeight) * 0.98)
      if (fitted < current - 0.001) {
        normalAutoFitScales.set(key, fitted)
        changed = true
      }
    }
    if (changed) void nextTick(scheduleShapeAutoFit)
  })
}

function onFontsLoaded() {
  autoFitDimensions.clear()
  normalAutoFitScales.clear()
  scheduleShapeAutoFit()
}
const chartPlot = (id: string) => chartVisuals.value.get(id)?.plot ?? { left: 88, top: 75, width: 850, height: 390 }
type RenderAnimation = PptxAnimation & { step: number; automatic: boolean; triggerSequence?: number; triggerStep?: number }
const triggerKey = (sequence: number, step: number) => `${sequence}:${step}`
const timelineStart = (action: RenderAnimation) => {
  const startedAt = action.triggerSequence === undefined || action.triggerStep === undefined
  ? action.automatic ? automaticStartedAt : stepStartedAt.get(action.step)
  : triggerStartedAt.get(triggerKey(action.triggerSequence, action.triggerStep))
  if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY || !action.mediaWaitForEndKeys?.length) return startedAt
  const endedAt = action.mediaWaitForEndKeys.map(key => completedMediaActions.get(key))
  if (endedAt.some(time => time === undefined)) return Infinity
  return Math.max(startedAt, ...endedAt.map(time => time! + (action.mediaWaitDelayMs || 0) - action.delayMs))
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
const timelineActions = computed<RenderAnimation[]>(() => [
  ...(props.slide.automaticAnimations || []).map(animation => ({ ...animation, step: -1, automatic: true })),
  ...(props.slide.animationSteps || []).flatMap((step, index) => step.map(animation => ({ ...animation, step: index, automatic: false }))),
  ...(props.slide.triggeredAnimations || []).flatMap((sequence, sequenceIndex) => sequence.steps.flatMap((step, stepIndex) => step.map(animation => ({ ...animation, step: -1, automatic: false, triggerSequence: sequenceIndex, triggerStep: stepIndex })))),
])
const mediaSlideCounts = computed(() => {
  const counts = new Map<string, number>()
  for (const action of timelineActions.value) {
    if (action.effect === 'media' && action.mediaSlideCount) counts.set(action.targetId, Math.max(action.mediaSlideCount, counts.get(action.targetId) || 0))
  }
  return counts
})

function cancelTimelineUpdates() {
  cancelAnimationFrame(animationFrame)
  animationFrame = 0
  if (animationTimer) clearTimeout(animationTimer)
  animationTimer = undefined
}

function scheduleTimelineUpdate() {
  cancelTimelineUpdates()
  if (props.thumbnail && !props.playbackPreview) return
  runTimedMediaActions()
  const now = animationClock.value
  if (reduceMotion.value) {
    const nextMediaStart = timelineActions.value.reduce((next, action) => {
      if (action.effect !== 'media') return next
      const startedAt = timelineStart(action)
      return startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY ? next : Math.min(next, startedAt + action.delayMs)
    }, Infinity)
    if (Number.isFinite(nextMediaStart) && nextMediaStart > now) {
      animationTimer = setTimeout(() => {
        animationTimer = undefined
        animationClock.value = performance.now()
        scheduleTimelineUpdate()
      }, Math.max(0, nextMediaStart - now))
    }
    return
  }
  let nextStart = Infinity
  let hasActiveAnimation = false
  for (const action of timelineActions.value) {
    const startedAt = timelineStart(action)
    if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY) continue
    const start = startedAt + action.delayMs
    const end = start + animationScheduleDuration(action)
    if (start > now) nextStart = Math.min(nextStart, start)
    else if (end > now) hasActiveAnimation = true
  }
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
  animationClock.value = performance.now()
  scheduleTimelineUpdate()
}

function beginSlideTimeline(step: number) {
  const now = performance.now()
  invalidMediaTrimIds.value = []
  invalidMediaFadeIds.value = []
  blockedMediaPlaybackIds.value = []
  executedMediaActions.clear()
  completedMediaActions.clear()
  for (const timer of mediaActionEndTimers.values()) clearTimeout(timer)
  mediaActionEndTimers.clear()
  automaticStartedAt = now
  stepStartedAt = new Map(Array.from({ length: step }, (_, index) => [index, Number.NEGATIVE_INFINITY]))
  triggerSteps.clear()
  triggerStartedAt.clear()
  triggerPlaybackByKey.clear()
  animationClock.value = now
  scheduleTimelineUpdate()
}

function syncTriggerPlayback() {
  if (props.triggerPlayback === undefined) return
  triggerSteps.clear()
  triggerStartedAt.clear()
  triggerPlaybackByKey.clear()
  for (const playback of props.triggerPlayback) {
    const sequence = props.slide.triggeredAnimations?.[playback.sequence]
    if (!sequence || !Number.isInteger(playback.step) || playback.step < 0 || playback.step >= sequence.steps.length || !Number.isFinite(playback.startedAt)) continue
    const key = triggerKey(playback.sequence, playback.step)
    triggerStartedAt.set(key, playback.startedAt - performance.timeOrigin)
    triggerPlaybackByKey.set(key, playback)
    triggerSteps.set(playback.sequence, Math.max(triggerSteps.get(playback.sequence) || 0, playback.step + 1))
  }
  refreshTimeline()
}

onMounted(() => {
  if (!stage.value) return
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
    beginSlideTimeline(step)
    void nextTick(scheduleShapeAutoFit)
    return
  }
  if (step > oldStep) {
    for (let previous = oldStep + 1; previous < step; previous++) stepStartedAt.set(previous, Number.NEGATIVE_INFINITY)
    if (step === oldStep + 1) runNavigationTriggeredAnimations('onNext')
    const now = performance.now()
    if (props.slide.animationSequence?.nextAction === 'seek' && oldStep > 0) {
      const previousStep = oldStep - 1
      const startedAt = stepStartedAt.get(previousStep)
      const naturalEnd = Math.max(0, ...(props.slide.animationSteps?.[previousStep] || []).map(action => action.delayMs + naturalAnimationDuration(action)))
      if (startedAt !== undefined && startedAt !== Number.NEGATIVE_INFINITY) stepStartedAt.set(previousStep, Math.min(startedAt, now - naturalEnd))
    }
    stepStartedAt.set(oldStep, now)
  } else if (step < oldStep) {
    if (step === oldStep - 1) runNavigationTriggeredAnimations('onPrev')
    for (const startedStep of stepStartedAt.keys()) {
      if (startedStep >= step) stepStartedAt.delete(startedStep)
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
  document.fonts.removeEventListener('loadingdone', onFontsLoaded)
  cancelAnimationFrame(autoFitFrame)
  motionPreference?.removeEventListener('change', onMotionPreferenceChange)
  cancelTimelineUpdates()
  for (const frame of mediaFadeFrames.values()) cancelAnimationFrame(frame)
  mediaFadeFrames.clear()
  for (const timer of mediaActionEndTimers.values()) clearTimeout(timer)
  mediaActionEndTimers.clear()
})

function onMotionPreferenceChange(event: MediaQueryListEvent) {
  reduceMotion.value = event.matches
  refreshTimeline()
}

function applyColorOpacity(color: string, opacity: number): string {
  if (opacity >= 1 || color === 'transparent') return color
  const hex = /^#([\da-f]{6})$/i.exec(color)
  const rgb = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)(?:\s*,\s*([\d.]+)\s*)?\)$/i.exec(color)
  const channels = hex
    ? [parseInt(hex[1]!.slice(0, 2), 16), parseInt(hex[1]!.slice(2, 4), 16), parseInt(hex[1]!.slice(4, 6), 16), 1]
    : rgb ? [Number(rgb[1]), Number(rgb[2]), Number(rgb[3]), rgb[4] === undefined ? 1 : Number(rgb[4])] : undefined
  if (!channels) return color
  return `rgba(${channels.slice(0, 3).join(', ')}, ${Math.max(0, Math.min(channels[3]! * opacity, 1))})`
}

function animatedRunColor(color: string | undefined, opacity: number): string | undefined {
  if (opacity >= 1) return color
  const value = color || 'currentColor'
  const adjusted = applyColorOpacity(value, opacity)
  return adjusted !== value || value === 'transparent'
    ? adjusted
    : `color-mix(in srgb, ${value} ${Math.max(0, Math.min(opacity, 1)) * 100}%, transparent)`
}

function elementStyle(element: PptxElement): CSSProperties {
  const fittedSize = autoFitDimensions.get(`${props.slide.id}:${element.id}`)
  const width = fittedSize?.width || element.width
  const height = fittedSize?.height || element.height
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
  const shadow = element.shadow
  const angle = (shadow?.angle || 0) * Math.PI / 180
  const shadowColor = animationColor(element, 'shadow.color') || shadow?.color
  const animatedFill = animationColor(element, 'fillcolor')
  const fill = animatedFill || element.fill
  const fillOpacity = animationOpacity(element, 'fill.opacity')
  const animatedStroke = animationColor(element, 'stroke.color')
  const stroke = applyColorOpacity(animatedStroke || element.stroke, animationOpacity(element, 'stroke.opacity'))
  const customGeometry = element.geometry === 'custom' && Boolean(element.customPaths?.length)
  const showStroke = element.kind === 'shape' && !customGeometry && element.strokeWidth > 0 && stroke !== 'transparent' && !stroke.includes('gradient')
  const transform = matrix
    ? [translations, `matrix(${matrix.slice(0, 4).join(',')},${matrix[4] * scale.value},${matrix[5] * scale.value})${hasScaleAnimation || hasRotationAnimation ? ` translate(${element.width * scale.value / 2}px,${element.height * scale.value / 2}px)${hasRotationAnimation ? ` rotate(${rotationDelta}deg)` : ''}${hasScaleAnimation ? ` scale(${scaleX},${scaleY})` : ''} translate(${-element.width * scale.value / 2}px,${-element.height * scale.value / 2}px)` : ''}`].filter(Boolean).join(' ')
    : [translations, rotation ? `rotate(${rotation}deg)` : '', flips, hasScaleAnimation ? `scale(${scaleX},${scaleY})` : ''].filter(Boolean).join(' ')
  return {
    left: `${(element.x / props.width) * 100}%`,
    top: `${(element.y / props.height) * 100}%`,
    width: `${(width / props.width) * 100}%`,
    height: `${(height / props.height) * 100}%`,
    overflow: customGeometry ? 'visible' : undefined,
    transform: transform || undefined,
    transformOrigin: matrix ? '0 0' : undefined,
    background: element.kind === 'shape' && !customGeometry ? applyColorOpacity(fill, fillOpacity) : undefined,
    borderColor: showStroke ? stroke : undefined,
    borderWidth: showStroke ? `${element.strokeWidth * scale.value}px` : undefined,
    borderStyle: showStroke ? cssLineStyle(element.strokeDash) : undefined,
    boxShadow: shadow && !customGeometry ? `${Math.cos(angle) * shadow.distance * scale.value}px ${Math.sin(angle) * shadow.distance * scale.value}px ${shadow.blur * scale.value}px ${applyColorOpacity(shadowColor || shadow.color, animationOpacity(element, 'shadow.opacity'))}` : undefined,
  }
}

function customGeometryFill(element: PptxElement, path: NonNullable<PptxElement['customPaths']>[number]): string {
  if (!path.fill) return 'none'
  const animatedFill = animationColor(element, 'fillcolor')
  if (!animatedFill && /^(?:linear|radial)-gradient\(/i.test(element.fill)) return `url(#${customGradientId(element, 'fill')})`
  return applyColorOpacity(animatedFill || element.fill, animationOpacity(element, 'fill.opacity'))
}

function customGeometryStroke(element: PptxElement, path: NonNullable<PptxElement['customPaths']>[number]): string {
  const animatedStroke = animationColor(element, 'stroke.color')
  if (path.stroke && element.strokeWidth > 0 && !animatedStroke && /^(?:linear|radial)-gradient\(/i.test(element.stroke)) return `url(#${customGradientId(element, 'stroke')})`
  const stroke = applyColorOpacity(animatedStroke || element.stroke, animationOpacity(element, 'stroke.opacity'))
  return path.stroke && element.strokeWidth > 0 && stroke !== 'transparent' && !stroke.includes('gradient') ? stroke : 'none'
}

function customGradientId(element: PptxElement, kind: 'fill' | 'stroke'): string {
  return `${customGradientScope}-${kind}-${element.id.replace(/[^a-zA-Z0-9_-]/g, '_')}`
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

function customSvgGradient(element: PptxElement, kind: 'fill' | 'stroke'): CustomSvgGradient | undefined {
  const css = kind === 'fill' ? element.fill : element.stroke
  const linear = /^linear-gradient\(\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+))deg\s*,([\s\S]*)\)$/i.exec(css)
  const radial = /^radial-gradient\(\s*(?:circle|ellipse)\s*,([\s\S]*)\)$/i.exec(css)
  if (!linear && !radial) return undefined
  const stops = splitGradientStops((linear ? linear[2] : radial![1])!).map((stop): CustomSvgGradientStop | undefined => {
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
      opacity: rgba && rgba[4] !== undefined ? Math.max(0, Math.min(Number(rgba[4]), 1)) : 1,
    }
  })
  if (stops.length < 2 || stops.some(stop => !stop)) return undefined

  if (radial) return { id: customGradientId(element, kind), kind: 'radial', cx: 500, cy: 500, r: 500, stops: stops as CustomSvgGradientStop[] }

  // DrawingML angles describe the direction of color change; CSS angles already retain that direction.
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill.angle?view=openxml-3.0.1
  // https://www.w3.org/TR/SVG2/pservers/
  const radians = Number(linear![1]) * Math.PI / 180
  const dx = Math.sin(radians), dy = -Math.cos(radians)
  const halfLength = (Math.abs(dx) + Math.abs(dy)) * 500
  return { id: customGradientId(element, kind), kind: 'linear', x1: 500 - dx * halfLength, y1: 500 - dy * halfLength, x2: 500 + dx * halfLength, y2: 500 + dy * halfLength, stops: stops as CustomSvgGradientStop[] }
}

function customGeometryGradients(element: PptxElement): CustomSvgGradient[] {
  return (['fill', 'stroke'] as const).flatMap(kind => {
    const gradient = customSvgGradient(element, kind)
    return gradient ? [gradient] : []
  })
}

function customGeometryStrokeWidth(element: PptxElement): string {
  return `${element.strokeWidth * scale.value}px`
}

function customGeometryStrokeStyle(element: PptxElement): CSSProperties {
  // DrawingML custom dash entries are length ratios of the rendered pen width.
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1
  // https://www.w3.org/TR/svg-strokes/
  const patterns: Record<string, number[]> = {
    dot: [0, 2], sysDot: [1, 2], dash: [3, 2], sysDash: [3, 3], lgDash: [6, 2],
    dashDot: [3, 2, 1, 2], sysDashDot: [3, 2, 1, 2], lgDashDot: [6, 2, 1, 2],
    lgDashDotDot: [6, 2, 1, 2, 1, 2], sysDashDotDot: [3, 2, 1, 2, 1, 2],
  }
  const pattern = Array.isArray(element.strokeDash) ? element.strokeDash : patterns[element.strokeDash || '']
  const width = Array.isArray(element.strokeDash) ? element.strokeWidth * scale.value : Math.max(element.strokeWidth * scale.value, 1)
  return {
    ...(pattern ? { strokeDasharray: pattern.map(length => `${length * width}px`).join(' ') } : {}),
    strokeLinecap: element.strokeDash === 'dot' || element.strokeDash === 'sysDot' ? 'round' : element.strokeCap || 'butt',
  }
}

function cssLineStyle(pattern?: PptxElement['strokeDash']): 'solid' | 'dashed' | 'dotted' {
  if (Array.isArray(pattern)) return 'dashed'
  if (pattern === 'dot' || pattern === 'sysDot') return 'dotted'
  return pattern && pattern !== 'solid' ? 'dashed' : 'solid'
}

function lineStrokeColor(element: PptxElement): string {
  return applyColorOpacity(animationColor(element, 'stroke.color') || (element.stroke === 'transparent' ? '#222' : element.stroke), animationOpacity(element, 'stroke.opacity'))
}

function lineRenderStyle(element: PptxElement): CSSProperties {
  return { width: `${Math.max(element.width * scale.value, 1)}px`, height: `${Math.max(element.height * scale.value, 1)}px`, color: lineStrokeColor(element) }
}

function lineEndId(element: PptxElement, side: 'head' | 'tail'): string {
  return `pptx-line-${element.id}-${side}`
}

function lineEndUrl(element: PptxElement, side: 'head' | 'tail'): string | undefined {
  const ending = side === 'head' ? element.lineHeadEnd : element.lineTailEnd
  return ending && ending.type !== 'none' ? `url(#${lineEndId(element, side)})` : undefined
}

function lineEndMarkerSize(element: PptxElement, side: 'head' | 'tail'): { width: number; height: number } {
  const ending = side === 'head' ? element.lineHeadEnd : element.lineTailEnd
  const size: Record<NonNullable<typeof ending>['width'], number> = { sm: 0.7, med: 1, lg: 1.4 }
  return { width: 6 * size[ending?.length || 'med'], height: 6 * size[ending?.width || 'med'] }
}

function lineEndPath(type: NonNullable<PptxElement['lineHeadEnd']>['type']): string {
  if (type === 'stealth') return 'M0 0L10 5L0 10L3.5 5Z'
  if (type === 'diamond') return 'M0 5L5 0L10 5L5 10Z'
  if (type === 'oval') return 'M5 0A5 5 0 1 1 5 10A5 5 0 1 1 5 0Z'
  if (type === 'arrow') return 'M0 0L10 5L0 10'
  return 'M0 0L10 5L0 10Z'
}

function customGeometryStyle(element: PptxElement): CSSProperties {
  const shadow = element.shadow
  const angle = (shadow?.angle || 0) * Math.PI / 180
  const shadowColor = animationColor(element, 'shadow.color') || shadow?.color
  return shadow ? {
    filter: `drop-shadow(${Math.cos(angle) * shadow.distance * scale.value}px ${Math.sin(angle) * shadow.distance * scale.value}px ${shadow.blur * scale.value}px ${applyColorOpacity(shadowColor || shadow.color, animationOpacity(element, 'shadow.opacity'))})`,
  } : {}
}

function imageStyle(element: PptxElement): CSSProperties {
  const crop = element.imageCrop
  if (!crop) return { width: '100%', height: '100%' }
  const width = 1 - crop.left - crop.right
  const height = 1 - crop.top - crop.bottom
  return {
    width: `${100 / width}%`,
    height: `${100 / height}%`,
    left: `${-crop.left / width * 100}%`,
    top: `${-crop.top / height * 100}%`,
  }
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
  const currentFrame = mediaFadeFrames.get(media)
  if (media.paused || media.ended || !element.mediaFade) {
    if (currentFrame) cancelAnimationFrame(currentFrame)
    mediaFadeFrames.delete(media)
  } else if (!currentFrame) {
    mediaFadeFrames.set(media, requestAnimationFrame(() => {
      mediaFadeFrames.delete(media)
      updateMediaFade(media, element)
    }))
  }
}

function updateMediaPlayback(event: Event, element: PptxElement) {
  applyMediaTrim(event, element)
  updateMediaFade(event.currentTarget as HTMLMediaElement, element)
}

function updateMediaFadeForEvent(event: Event, element: PptxElement) {
  updateMediaFade(event.currentTarget as HTMLMediaElement, element)
}

function updateMediaOnPlay(event: Event, element: PptxElement) {
  blockedMediaPlaybackIds.value = blockedMediaPlaybackIds.value.filter(id => id !== element.id)
  updateMediaPlayback(event, element)
}

function finishTimedMedia(element: PptxElement, refresh = true) {
  const endedAt = performance.now()
  let changed = false
  const media = Array.from(stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || []).find(item => item.dataset.mediaId === element.id)
  if (media) delete media.dataset.mediaActionEndAt
  timelineActions.value.forEach((action, index) => {
    if (action.effect !== 'media' || action.targetId !== element.id || !action.mediaActionKey || !executedMediaActions.has(index)) return
    completedMediaActions.set(action.mediaActionKey, endedAt)
    const timer = mediaActionEndTimers.get(action.mediaActionKey)
    if (timer) clearTimeout(timer)
    mediaActionEndTimers.delete(action.mediaActionKey)
    changed = true
  })
  if (changed && refresh) refreshTimeline()
}

function runTimedMediaActions() {
  timelineActions.value.forEach((action, index) => {
    if (action.effect !== 'media' || !action.mediaCommand || executedMediaActions.has(index)) return
    const startedAt = timelineStart(action)
    if (startedAt === undefined || startedAt === Number.NEGATIVE_INFINITY || animationClock.value < startedAt + action.delayMs) return
    executedMediaActions.add(index)
    const element = props.slide.elements.find(item => item.id === action.targetId)
    const media = Array.from(stage.value?.querySelectorAll<HTMLMediaElement>('[data-media-id]') || []).find(item => item.dataset.mediaId === action.targetId)
    if (!element || !media) return
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
      void media.play().catch(error => {
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
        media.dataset.mediaActionEndAt = String(performance.now() + action.mediaDurationMs)
        mediaActionEndTimers.set(action.mediaActionKey, setTimeout(() => {
          delete media.dataset.mediaActionEndAt
          media.pause()
          finishTimedMedia(element)
        }, action.mediaDurationMs))
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
  if (action.durationMs === 0) return action.reversePlayback ? 0 : 1
  const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1)
  const totalDuration = animationDuration(action)
  if (totalDuration === 0) return 0
  if (!Number.isFinite(totalDuration)) return 0
  const remainder = totalDuration % cycleDuration
  let progress: number
  if (!action.autoReverse) progress = remainder === 0 ? 1 : easedProgress(action, remainder / action.durationMs)
  else if (remainder === 0) progress = 0
  else if (remainder <= action.durationMs) progress = easedProgress(action, remainder / action.durationMs)
  else progress = 1 - easedProgress(action, (remainder - action.durationMs) / action.durationMs)
  return action.reversePlayback ? 1 - progress : progress
}

function animationFillExpired(action: RenderAnimation, startedAt: number | undefined, iterationIndex = 0): boolean {
  if (action.fillMode === 'remove') {
    if (startedAt === undefined) return action.triggerSequence === undefined && !action.automatic && action.step < props.animationStep
    if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return true
    const stagger = action.iteration ? action.iteration.intervalMs * iterationIndex : 0
    return action.timingWarp
      ? timingWarpChildElapsed(action, startedAt, stagger) >= animationDuration(action)
      : animationClock.value >= startedAt + action.delayMs + stagger + animationDuration(action)
  }
  if (action.fillMode !== 'freeze' || startedAt === undefined) return false
  const actionStart = startedAt + action.delayMs
  return (animationsByTarget.value.get(action.targetId) || []).some((sibling) => {
    if (sibling.sequenceKey !== action.sequenceKey || (sibling.sequenceOrder ?? -1) <= (action.sequenceOrder ?? -1)) return false
    const siblingStart = timelineStart(sibling)
    if (siblingStart === undefined) return false
    const start = siblingStart + sibling.delayMs
    return start >= actionStart && start <= animationClock.value
  })
}

function timingWarpChildElapsed(action: RenderAnimation, startedAt: number, stagger = 0): number {
  const warp = action.timingWarp!
  const elapsedWallMs = animationClock.value - startedAt - warp.groupOffsetMs
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
    - warp.childOffsetMs - stagger
}

function animationProgress(action: RenderAnimation, iterationIndex?: number): number | null {
  if (props.thumbnail && !props.playbackPreview) return null
  if (action.iteration && (iterationIndex === undefined || iterationIndex < 0)) return null
  const startedAt = timelineStart(action)
  const stagger = action.iteration ? action.iteration.intervalMs * (iterationIndex || 0) : 0
  if (animationFillExpired(action, startedAt, iterationIndex || 0)) return -2
  if (startedAt === undefined) return action.triggerSequence === undefined && !action.automatic && action.step < props.animationStep ? finalAnimationProgress(action) : null
  if (startedAt === Infinity) return -1
  if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return finalAnimationProgress(action)
  const elapsed = action.timingWarp
    ? timingWarpChildElapsed(action, startedAt, stagger)
    : animationClock.value - startedAt - action.delayMs - stagger
  if (elapsed < 0) return -1
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

function animationScale(element: PptxElement): [number, number] {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'scale' && !action.paragraphRange && !action.characterRange) || []
  let value: [number, number] = [1, 1]
  for (const action of actions) {
    const progress = animationProgress(action)
    if (progress === -2) {
      value = [1, 1]
      continue
    }
    if (progress === null || progress < 0 || !action.scaleFrom || !action.scaleTo) continue
    value = [
      action.scaleFrom[0] + (action.scaleTo[0] - action.scaleFrom[0]) * progress,
      action.scaleFrom[1] + (action.scaleTo[1] - action.scaleFrom[1]) * progress,
    ]
  }
  return value
}

function animationRotation(element: PptxElement): number {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'rotation' && !action.paragraphRange && !action.characterRange) || []
  let value = element.rotation
  for (const action of actions) {
    const progress = animationProgress(action)
    if (progress === -2) {
      value = element.rotation
      continue
    }
    if (progress === null || progress < 0 || action.rotationFrom === undefined || action.rotationTo === undefined) continue
    const rotation = action.rotationFrom + (action.rotationTo - action.rotationFrom) * progress
    value = action.rotationRelative ? value + rotation : rotation
  }
  return value
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
  if (nextIndex < 0) return keyframes.at(-1)!.value
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
    const hasTextRange = Boolean(action.paragraphRange || action.characterRange)
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
    if (progress === -2) {
      opacity = 1
      continue
    }
    if (progress === null || progress < 0 || action.opacityFrom === undefined || action.opacityTo === undefined) continue
    const keyframes = action.opacityKeyframes
    if (keyframes?.length) opacity = animationKeyframeValue(keyframes, action.opacityKeyframeMode, progress)
    else opacity = action.opacityFrom + (action.opacityTo - action.opacityFrom) * progress
  }
  return opacity
}

function animationColor(element: PptxElement, property: NonNullable<PptxAnimation['colorProperty']>, paragraphIndex?: number, characterStart?: number, characterEnd?: number): string | undefined {
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
    const from = action.colorFrom, to = action.colorTo
    if (progress === null || progress < 0 || !from || !to) continue
    let channels: [number, number, number, number]
    if (action.colorSpace === 'hsl') {
      const hueDelta = ((to[0] - from[0]) % 1 + 1) % 1
      const huePath = action.colorDirection === 'ccw' && hueDelta ? hueDelta - 1 : hueDelta
      const [r, g, b] = hslToRgb(from[0] + huePath * progress, from[1] + (to[1] - from[1]) * progress, from[2] + (to[2] - from[2]) * progress)
      channels = [r, g, b, from[3] + (to[3] - from[3]) * progress]
    } else {
      channels = [0, 1, 2, 3].map(index => from[index]! + (to[index]! - from[index]!) * progress) as [number, number, number, number]
    }
    color = `rgba(${channels.slice(0, 3).map(channel => Math.round(Math.max(0, Math.min(channel, 255)))).join(', ')}, ${Math.max(0, Math.min(channels[3], 1))})`
  }
  return color
}

function textColorAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties {
  const color = animationColor(element, 'style.color', paragraphIndex, characterStart, characterEnd)
  if (!color) return {}
  return { color: animatedRunColor(color, animationOpacity(element, 'fill.opacity', paragraphIndex, characterStart, characterEnd)) }
}

function textRangeOpacityAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties {
  const opacity = animationOpacity(element, 'style.opacity', paragraphIndex, characterStart, characterEnd)
  return opacity === 1 ? {} : { opacity }
}

function textRunAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number, run: PptxElement['paragraphs'][number]['runs'][number], fontScale = element.textFontScale ?? 1, script: ScriptFont = 'latin'): CSSProperties {
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
    factor = action.fontSizeKeyframes?.length
      ? animationKeyframeValue(action.fontSizeKeyframes, action.fontSizeKeyframeMode, progress)
      : action.fontSizeFrom + (action.fontSizeTo - action.fontSizeFrom) * progress
  }
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
      weight = run.bold ? 700 : 400
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.fontWeightReset) {
      weight = run.bold ? 700 : 400
      continue
    }
    if (action.fontWeightFrom === undefined || action.fontWeightTo === undefined) continue
    weight = action.fontWeightFormulaSamples?.length
      ? animationKeyframeValue(action.fontWeightFormulaSamples, 'lin', progress)
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
    if (progress === -2 && action.textStyleProperty === 'textTransform') {
      verticalAlign = run.baselineOffsetEm === undefined ? 'baseline' : `${run.baselineOffsetEm}em`
      verticalAlignChanged = true
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
    if (action.textStyleProperty === 'fontStyle') italic = value as boolean
    else if (action.textStyleProperty === 'fontFamily') fontFamily = cssFontFamily(String(value))
    else if (action.textStyleProperty === 'underline') underline = value as boolean
    else if (action.textStyleProperty === 'lineThrough') lineThrough = value as boolean
    else if (action.textStyleProperty === 'textTransform') {
      verticalAlign = value === 'sub' || value === 'super' ? value : 'baseline'
      verticalAlignChanged = true
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
  const textDecorationLine = [
    (underline ?? run.underline) ? 'underline' : undefined,
    (lineThrough ?? run.lineThrough) ? 'line-through' : undefined,
  ].filter(Boolean).join(' ')
  const embossOffset = Math.max(0.5, (run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value / 12)
  return {
    ...(factor === 1 ? {} : { fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value * factor}px` }),
    ...(weight === undefined ? {} : { fontWeight: weight }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(italic === undefined ? {} : { fontStyle: italic ? 'italic' : 'normal' }),
    ...(underline === undefined && lineThrough === undefined ? {} : { textDecorationLine: textDecorationLine || 'none' }),
    ...(verticalAlignChanged ? { verticalAlign } : {}),
    ...(textOutlineChanged ? { WebkitTextStroke: textOutline } : {}),
    ...(textShadowChanged ? { textShadow } : {}),
    ...(textEmbossChanged ? { filter: textEmbossEnabled ? `drop-shadow(-${embossOffset}px -${embossOffset}px 0 rgba(255, 255, 255, 0.65)) drop-shadow(${embossOffset}px ${embossOffset}px 0 rgba(0, 0, 0, 0.65))` : 'none' } : {}),
  }
}

function textTransformAnimationStyle(element: PptxElement, paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => {
    if (!['scale', 'rotation', 'motion', 'slide'].includes(action.effect) || !(action.paragraphRange || action.characterRange)) return false
    return (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
      && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
  }) || []
  if (!actions.length) return {}
  let scaling: [number, number] = [1, 1]
  let rotation = 0
  const [motionX, motionY] = animationMotion(element, paragraphIndex, characterStart, characterEnd)
  for (const action of actions) {
    const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd))
    if (progress === -2) {
      if (action.effect === 'scale') scaling = [1, 1]
      else rotation = 0
      continue
    }
    if (progress === null || progress < 0) continue
    if (action.effect === 'scale' && action.scaleFrom && action.scaleTo) {
      scaling = [
        action.scaleFrom[0] + (action.scaleTo[0] - action.scaleFrom[0]) * progress,
        action.scaleFrom[1] + (action.scaleTo[1] - action.scaleFrom[1]) * progress,
      ]
    } else if (action.effect === 'rotation' && action.rotationFrom !== undefined && action.rotationTo !== undefined) {
      const angle = action.rotationFrom + (action.rotationTo - action.rotationFrom) * progress
      rotation = action.rotationRelative ? rotation + angle : angle
    }
  }
  const transforms = [
    motionX || motionY ? `translate(${motionX * scale.value}px, ${motionY * scale.value}px)` : '',
    scaling[0] !== 1 || scaling[1] !== 1 ? `scale(${scaling[0]}, ${scaling[1]})` : '',
    rotation ? `rotate(${rotation}deg)` : '',
  ].filter(Boolean)
  return { display: 'inline-block', transform: transforms.join(' '), transformOrigin: 'center center' }
}

function animationMotion(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): [number, number] {
  const textRange = paragraphIndex !== undefined && characterStart !== undefined && characterEnd !== undefined
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'motion'
    && (textRange
      ? Boolean(action.paragraphRange || action.characterRange)
        && (!action.paragraphRange || paragraphIndex! >= action.paragraphRange.start && paragraphIndex! <= action.paragraphRange.end)
        && (!action.characterRange || characterStart! < action.characterRange.end && characterEnd! > action.characterRange.start)
      : !action.paragraphRange && !action.characterRange)) || []
  let value: [number, number] = [0, 0]
  for (const action of actions) {
    const progress = animationProgress(action, textRange ? iterationIndexFor(action, characterStart!, characterEnd!) : undefined)
    if (progress === -2) {
      value = [0, 0]
      continue
    }
    const samples = action.motionPathSamples
    if (progress !== null && progress >= 0 && samples?.length) {
      const samplePosition = Math.min(samples.length - 1, progress * (samples.length - 1))
      const from = samples[Math.floor(samplePosition)]!
      const to = samples[Math.ceil(samplePosition)]!
      const ratio = samplePosition - Math.floor(samplePosition)
      value = [from.x + (to.x - from.x) * ratio, from.y + (to.y - from.y) * ratio]
      continue
    }
    const path = action.motionPath
    if (progress === null || progress < 0 || !path?.length) continue
    if (path.length === 1) {
      const point = path[0]
      if (point) value = [point.x, point.y]
      continue
    }
    const distance = (action.motionPathLength || 0) * progress
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
      value = [from.x, from.y]
      continue
    }
    const span = to.distance - from.distance
    const ratio = span ? (distance - from.distance) / span : 0
    value = [from.x + (to.x - from.x) * ratio, from.y + (to.y - from.y) * ratio]
  }
  return value
}

function animationSlide(element: PptxElement): [number, number] {
  let value: [number, number] = [0, 0]
  for (const action of animationsByTarget.value.get(element.id)?.filter(item => item.effect === 'slide') || []) {
    const progress = animationProgress(action)
    if (progress === -2) { value = [0, 0]; continue }
    if (progress === null) continue
    const phase = progress < 0 ? 0 : progress
    const distance = action.direction === 'in' ? 1 - phase : phase
    value = action.slideFrom === 'left' ? [-props.width * distance, 0]
      : action.slideFrom === 'right' ? [props.width * distance, 0]
        : action.slideFrom === 'top' ? [0, -props.height * distance]
          : action.slideFrom === 'bottom' ? [0, props.height * distance]
            : [0, 0]
  }
  return value
}

function animationWipe(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    if (action.effect !== 'wipe') return false
    if (characterStart !== undefined && characterEnd !== undefined) {
      return Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
    }
    return paragraphIndex === undefined
      ? !action.paragraphRange && !action.characterRange
      : Boolean(action.paragraphRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
  }) || []
  const first = actions[0]
  if (!first?.wipeDirection) return {}
  let progress = first.direction === 'in' ? 0 : 100
  let direction = first.wipeDirection
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
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
  const cssDirection = direction === 'up' ? 'top' : direction === 'down' ? 'bottom' : direction
  const mask = `linear-gradient(to ${cssDirection}, #000 var(--wipe-progress), transparent var(--wipe-progress))`
  return { '--wipe-progress': `${progress}%`, maskImage: mask, WebkitMaskImage: mask } as CSSProperties
}

function animationBlinds(element: PptxElement, paragraphIndex?: number, characterStart?: number, characterEnd?: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => {
    if (action.effect !== 'blinds') return false
    if (paragraphIndex !== undefined) {
      return Boolean(action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
    }
    if (characterStart !== undefined && characterEnd !== undefined) {
      return Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
    }
    return !action.characterRange && !action.paragraphRange
  }) || []
  const first = actions[0]
  if (!first?.blindsOrientation) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let orientation = first.blindsOrientation
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
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
    if (paragraphIndex !== undefined) return Boolean(action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end)
    if (characterStart !== undefined && characterEnd !== undefined) return Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start)
    return !action.paragraphRange && !action.characterRange
  }) || []
  const first = actions[0]
  if (!first?.checkerOrientation) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let orientation = first.checkerOrientation
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
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
  const width = element.width * scale.value
  const height = element.height * scale.value
  const amount = Math.max(0, Math.min(progress, 1))
  if (paragraphIndex === undefined && characterStart === undefined) return { clipPath: gridTransitionClipPath(width, height, orientation, amount, true) }
  const mask = gridTransitionMaskImage(width, height, orientation, amount, true)
  return {
    maskImage: mask, WebkitMaskImage: mask,
    maskSize: '100% 100%', WebkitMaskSize: '100% 100%',
    maskRepeat: 'no-repeat', WebkitMaskRepeat: 'no-repeat',
  }
}

function animationShapeMask(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'shape') || []
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
    const phase = animationProgress(action)
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
  return { clipPath: shapeEffectClipPath(current.shape, element.width * scale.value, element.height * scale.value, current.progress, current.direction) }
}

function animationDissolve(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'dissolve') || []
  const first = actions[0]
  if (!first) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
    if (phase === -2) {
      reset = true
      continue
    }
    if (phase === null || phase < 0) continue
    reset = false
    progress = action.direction === 'in' ? phase : 1 - phase
  }
  if (reset) return {}
  return { clipPath: dissolveClipPath(element.width * scale.value, element.height * scale.value, progress, element.id) }
}

function animationWheel(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'wheel') || []
  const first = actions[0]
  if (!first?.wheelSpokes) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let spokes = first.wheelSpokes
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
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
  return { clipPath: wheelClipPath(element.width * scale.value, element.height * scale.value, spokes, progress) }
}

function animationRandomBars(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'randomBars') || []
  const first = actions[0]
  if (!first?.randomBarOrientation) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let orientation = first.randomBarOrientation
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
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
  return { clipPath: randomBarClipPath(element.width * scale.value, element.height * scale.value, orientation, progress, element.id) }
}

function animationStrips(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'strips') || []
  const first = actions[0]
  if (!first?.stripsDirection) return {}
  let progress = first.direction === 'in' ? 0 : 1
  let direction = first.stripsDirection
  let reset = false
  for (const action of actions) {
    const phase = animationProgress(action)
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
  return { clipPath: stripsClipPath(element.width * scale.value, element.height * scale.value, direction, progress) }
}

function animationBarn(element: PptxElement): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter(action => action.effect === 'barn') || []
  const first = actions[0]
  if (!first?.barnOrientation || !first.barnMotion) return {}
  let orientation = first.barnOrientation
  let motion = first.barnMotion
  let progress = 0
  let reset = false
  for (const action of actions) {
    if (!action.barnOrientation || !action.barnMotion) continue
    const phase = animationProgress(action)
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
  const visibilityActions = actions.filter(action => ['appear', 'fade', 'wipe', 'strips', 'barn', 'slide'].includes(action.effect))
  const first = visibilityActions[0]
  if (!first) return {}
  let opacity = first.effect === 'wipe' || first.effect === 'strips' || first.effect === 'barn' || first.effect === 'slide' || first.direction === 'out' ? 1 : 0
  let visibility: 'visible' | 'hidden' = first.direction === 'out' ? 'visible' : 'hidden'
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
    } else if (action.effect === 'wipe' || action.effect === 'strips' || action.effect === 'barn') {
      visibility = action.direction === 'in'
        ? progress > 0 ? 'visible' : 'hidden'
        : progress >= 1 ? 'hidden' : 'visible'
      opacity = 1
    } else if (action.effect === 'slide') {
      visibility = action.direction === 'in' || progress < 1 ? 'visible' : 'hidden'
      opacity = 1
    } else {
      const easedProgress = 1 - (1 - progress) ** 2
      opacity = action.direction === 'in' ? easedProgress : 1 - easedProgress
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
  const animatedStyle = animationCss(animationsByTarget.value.get(element.id)?.filter(action => !action.paragraphRange && !action.characterRange && !action.iteration))
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
    return range && !action.iteration && index >= range.start && index <= range.end
  })
  return { ...animationCss(actions), ...animationWipe(element, index), ...animationBlinds(element, index), ...animationChecker(element, index) }
}

function characterAnimationStyle(element: PptxElement, start: number, end: number): CSSProperties {
  const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
    const range = action.characterRange
    const overlaps = range && start < range.end && end > range.start
    if (action.iteration) return Boolean(overlaps || !range) && (iterationIndexFor(action, start, end) ?? -1) >= 0
    return Boolean(overlaps)
  })
  return { ...animationCss(actions, start, end), ...animationWipe(element, undefined, start, end), ...animationBlinds(element, undefined, start, end), ...animationChecker(element, undefined, start, end) }
}

function hasTrigger(elementId: string) {
  return !props.thumbnail && Boolean(props.slide.triggeredAnimations?.some(sequence => sequence.trigger.type === 'shape' && sequence.trigger.id === elementId))
}

function startTriggeredAnimations(indices: number[]) {
  if (props.triggerPlayback !== undefined) return
  const sequences = props.slide.triggeredAnimations || []
  if (!indices.length) return
  const now = performance.now()
  for (const index of indices) {
    const sequence = sequences[index]
    if (!sequence) continue
    const count = sequence.steps.length
    if (!count) continue
    let step = triggerSteps.get(index) || 0
    if (step >= count) {
      for (let previous = 0; previous < count; previous++) {
        const key = triggerKey(index, previous)
        triggerStartedAt.delete(key)
        triggerPlaybackByKey.delete(key)
      }
      timelineActions.value.forEach((action, actionIndex) => {
        if (action.triggerSequence === index) executedMediaActions.delete(actionIndex)
      })
      step = 0
    }
    const key = triggerKey(index, step)
    const playback = { sequence: index, step, startedAt: performance.timeOrigin + now }
    triggerStartedAt.set(key, now)
    triggerPlaybackByKey.set(key, playback)
    triggerSteps.set(index, step + 1)
  }
  refreshTimeline()
  emit('trigger-playback', [...triggerPlaybackByKey.values()].sort((a, b) => a.sequence - b.sequence || a.step - b.step))
}

function runTriggeredAnimations(event: Event, elementId: string, triggerEvent: 'onClick' | 'onDblClick' | 'onMouseOver' | 'onMouseOut') {
  if (props.thumbnail || props.triggerPlayback !== undefined) return
  const sequences = props.slide.triggeredAnimations || []
  const matching = sequences.flatMap((sequence, index) => sequence.trigger.type === 'shape' && sequence.trigger.id === elementId && sequence.trigger.event === triggerEvent ? [index] : [])
  if (!matching.length) {
    if (sequences.some(sequence => sequence.trigger.type === 'shape' && sequence.trigger.id === elementId)) event.stopPropagation()
    return
  }
  if (triggerEvent === 'onClick' || triggerEvent === 'onDblClick') event.stopPropagation()
  startTriggeredAnimations(matching)
}

function runNavigationTriggeredAnimations(triggerEvent: 'onNext' | 'onPrev') {
  if (props.thumbnail || props.triggerPlayback !== undefined) return
  const matching = (props.slide.triggeredAnimations || []).flatMap((sequence, index) => sequence.trigger.type === 'slide' && sequence.trigger.event === triggerEvent ? [index] : [])
  startTriggeredAnimations(matching)
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
  const sequence = props.slide.triggeredAnimations?.find(item => item.trigger.type === 'shape' && item.trigger.id === elementId)
  runTriggeredAnimations(event, elementId, sequence?.trigger.type === 'shape' ? sequence.trigger.event : 'onClick')
}

function usesRightToLeftColumns(element: PptxElement): boolean {
  return element.textColumnsRightToLeft === true
    && (element.textColumnCount ?? 1) > 1
    && (!element.textOrientation || element.textOrientation === 'horz')
}

function textFrameStyle(element: PptxElement): CSSProperties {
  const verticalText = element.textOrientation === 'vert' || element.textOrientation === 'vert270'
  const eastAsianVerticalText = element.textOrientation === 'eaVert'
  const wordArtVerticalText = element.textOrientation === 'wordArtVert'
  const wordArtVerticalRtlText = element.textOrientation === 'wordArtVertRtl'
  const mongolianVerticalText = element.textOrientation === 'mongolianVert'
  const columnCount = element.textColumnCount ?? 1
  const rect = element.customTextRect
  return {
    position: rect ? 'absolute' : undefined,
    left: rect ? `${rect.left * 100}%` : undefined,
    top: rect ? `${rect.top * 100}%` : undefined,
    width: rect ? `${(rect.right - rect.left) * 100}%` : undefined,
    height: rect ? `${(rect.bottom - rect.top) * 100}%` : undefined,
    boxSizing: rect ? 'border-box' : undefined,
    display: columnCount > 1 ? 'block' : undefined,
    columnCount: columnCount > 1 ? columnCount : undefined,
    columnGap: columnCount > 1 ? `${(element.textColumnSpacing ?? 0) * scale.value}px` : undefined,
    columnFill: columnCount > 1 ? 'auto' : undefined,
    padding: `${element.margins.top * scale.value}px ${element.margins.right * scale.value}px ${element.margins.bottom * scale.value}px ${element.margins.left * scale.value}px`,
    justifyContent: element.verticalAlign === 'middle' ? 'center' : element.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start',
    whiteSpace: element.textWrap === 'none' ? 'nowrap' : undefined,
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textverticalvalues?view=openxml-3.0.1
    // https://www.w3.org/TR/css-writing-modes-3/#block-flow
    writingMode: element.textOrientation === 'vert270' || mongolianVerticalText ? 'vertical-lr' : verticalText || eastAsianVerticalText || wordArtVerticalText ? 'vertical-rl' : undefined,
    textOrientation: verticalText ? 'sideways' : eastAsianVerticalText || wordArtVerticalText || mongolianVerticalText ? 'upright' : undefined,
    direction: element.textOrientation === 'vert270' ? 'rtl' : undefined,
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

function geometryStyle(geometry: string): CSSProperties {
  if (geometry === 'ellipse') return { borderRadius: '50%' }
  if (geometry === 'roundRect') return { borderRadius: '10%' }
  const clipPaths: Record<string, string> = {
    triangle: 'polygon(50% 0, 100% 100%, 0 100%)',
    rtTriangle: 'polygon(0 0, 100% 100%, 0 100%)',
    diamond: 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)',
    parallelogram: 'polygon(22% 0, 100% 0, 78% 100%, 0 100%)',
    hexagon: 'polygon(25% 0, 75% 0, 100% 50%, 75% 100%, 25% 100%, 0 50%)',
    chevron: 'polygon(0 0, 72% 0, 100% 50%, 72% 100%, 0 100%, 28% 50%)',
  }
  return clipPaths[geometry] ? { clipPath: clipPaths[geometry] } : {}
}

type ScriptFont = 'latin' | 'eastAsia' | 'complexScript'

type TextSegment = { text: string; script: ScriptFont; start: number; end: number }

function cssFontFamily(fontFamily: string | undefined): string {
  const safeFontFamily = (fontFamily || '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim().slice(0, 128)
  return safeFontFamily ? `'${safeFontFamily}', Arial, sans-serif` : 'Arial, sans-serif'
}

function runTextShadowStyle(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties['textShadow'] {
  const shadow = run.shadow
  if (!shadow) return undefined
  const angle = (shadow.angle || 0) * Math.PI / 180
  const shadowColor = animationColor(element, 'shadow.color', paragraphIndex, characterStart, characterEnd) || shadow.color
  return `${Math.cos(angle) * shadow.distance * scale.value}px ${Math.sin(angle) * shadow.distance * scale.value}px ${shadow.blur * scale.value}px ${animatedRunColor(shadowColor, animationOpacity(element, 'shadow.opacity', paragraphIndex, characterStart, characterEnd))}`
}

function runStyle(element: PptxElement, run: PptxElement['paragraphs'][number]['runs'][number], script: ScriptFont, fontScale: number, paragraphIndex: number, characterStart: number, characterEnd: number): CSSProperties {
  const requestedFont = script === 'eastAsia' ? run.fontFamilyEastAsia : script === 'complexScript' ? run.fontFamilyComplexScript : run.fontFamily
  const fillOpacity = animationOpacity(element, 'fill.opacity', paragraphIndex, characterStart, characterEnd)
  const strokeOpacity = animationOpacity(element, 'stroke.opacity', paragraphIndex, characterStart, characterEnd)
  const textDecorationLine = [run.underline ? 'underline' : undefined, run.lineThrough ? 'line-through' : undefined].filter(Boolean).join(' ')
  const textDecorationStyle = run.underline && run.lineThrough
    ? (run.underlineStyle ?? (run.lineThroughStyle === 'double' ? 'solid' : run.lineThroughStyle))
    : run.underlineStyle ?? run.lineThroughStyle
  return {
    fontFamily: cssFontFamily(requestedFont || run.fontFamily),
    fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value}px`,
    letterSpacing: run.letterSpacingPt === undefined ? undefined : `${run.letterSpacingPt * fontScale * EMU_PER_POINT * scale.value}px`,
    verticalAlign: run.baselineOffsetEm === undefined ? undefined : `${run.baselineOffsetEm}em`,
    color: animatedRunColor(run.color, fillOpacity),
    WebkitTextStroke: run.outlineColor !== undefined && run.outlineWidth !== undefined ? `${run.outlineWidth * scale.value}px ${animatedRunColor(run.outlineColor, strokeOpacity)}` : undefined,
    fontWeight: run.bold ? 700 : undefined,
    fontStyle: run.italic ? 'italic' : undefined,
    textDecorationLine: textDecorationLine || undefined,
    textDecorationStyle,
    textShadow: runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd),
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
      const script: ScriptFont = EAST_ASIAN_SCRIPT.test(character) || /[\u3000-\u303f\uff00-\uffef]/u.test(character)
        ? 'eastAsia'
        : COMPLEX_SCRIPT.test(character)
          ? 'complexScript'
          : /\p{Mark}/u.test(character)
            ? previousScript
            : 'latin'
      if (segmentScript && script !== segmentScript) {
        segments.push({ text: run.text.slice(segmentStart, offset), script: segmentScript, start: runStart + segmentStart, end: runStart + offset })
        segmentStart = offset
      }
      segmentScript = script
      previousScript = script
      offset += character.length
    }
    if (segmentScript && offset > segmentStart) segments.push({ text: run.text.slice(segmentStart), script: segmentScript, start: runStart + segmentStart, end: runStart + offset })
    return segments
  }

  const characters: { text: string; script: ScriptFont; start: number; end: number }[] = []
  const cuts = new Set([0, run.text.length])
  let previousScript: ScriptFont = 'latin'
  let offset = 0
  for (const character of run.text) {
    const script: ScriptFont = EAST_ASIAN_SCRIPT.test(character) || /[\u3000-\u303f\uff00-\uffef]/u.test(character)
      ? 'eastAsia'
      : COMPLEX_SCRIPT.test(character)
        ? 'complexScript'
        : /\p{Mark}/u.test(character)
          ? previousScript
          : 'latin'
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
  return segments
}

function alphabeticNumber(value: number): string {
  let number = Math.max(1, value)
  let result = ''
  while (number > 0) {
    number--
    result = String.fromCharCode(97 + number % 26) + result
    number = Math.floor(number / 26)
  }
  return result
}

function romanNumber(value: number): string {
  let number = Math.max(1, value)
  let result = ''
  for (const [amount, numeral] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']] as [number, string][]) {
    while (number >= amount) {
      result += numeral
      number -= amount
    }
  }
  return result
}

function numberedMarker(value: number, type: string): string {
  if (type === 'arabicPlain') return String(value)
  let label = String(value)
  if (type.startsWith('alpha')) {
    label = alphabeticNumber(value)
    if (type.startsWith('alphaUc')) label = label.toUpperCase()
  } else if (type.startsWith('roman')) {
    label = romanNumber(value)
    if (type.startsWith('romanUc')) label = label.toUpperCase()
  }
  if (type.endsWith('ParenBoth')) return `(${label})`
  if (type.endsWith('ParenR')) return `${label})`
  return `${label}.`
}

function paragraphEntries(paragraphs: PptxParagraph[]): { paragraph: PptxParagraph; marker: string }[] {
  const counters = new Map<string, number>()
  return paragraphs.map((paragraph) => {
    const bullet = paragraph.bullet
    if (!bullet) {
      counters.clear()
      return { paragraph, marker: '' }
    }
    if (bullet.kind === 'char') {
      counters.clear()
      return { paragraph, marker: bullet.char }
    }
    const key = `${bullet.type}:${paragraph.level || 0}`
    const number = counters.get(key) ?? bullet.startAt ?? 1
    counters.set(key, number + 1)
    return { paragraph, marker: numberedMarker(number, bullet.type) }
  })
}

function cssSpacing(spacing: PptxParagraph['lineSpacing']): string | undefined {
  if (!spacing) return undefined
  return spacing.unit === 'emu' ? `${spacing.value * scale.value}px` : `${spacing.value}em`
}

function paragraphStyle(paragraph: PptxParagraph, fontScale = 1, lineSpaceReduction = 0): CSSProperties {
  const lineSpacing = paragraph.lineSpacing
  const fontSizePt = paragraph.runs.reduce((size, run) => Math.max(size, run.fontSizePt || 18), 0) || 18
  return {
    fontSize: `${fontSizePt * fontScale * EMU_PER_POINT * scale.value}px`,
    textAlign: paragraph.align,
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

function tableCellTextStyle(cell: PptxTableCell): CSSProperties {
  const columnCount = cell.textColumnCount ?? 1
  return columnCount > 1 ? {
    height: '100%',
    boxSizing: 'border-box',
    columnCount,
    columnGap: `${(cell.textColumnSpacing ?? 0) * scale.value}px`,
    columnFill: 'auto',
  } : {}
}
</script>

<template>
  <div
    ref="stage"
    class="slide-frame"
    :class="{ 'is-thumbnail': thumbnail }"
    :style="frameStyle"
    :role="hasInteractiveMedia ? 'group' : 'img'"
    :aria-label="`Slide ${number}`"
  >
    <div
      v-for="element in slide.elements"
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
      :aria-label="element.hyperlink?.tooltip || (hasTrigger(element.id) ? `Play animation from ${element.name}` : undefined)"
      :data-pptx-hyperlink="canActivateHyperlink(element.hyperlink) ? JSON.stringify(element.hyperlink) : undefined"
      :style="[elementStyle(element), geometryStyle(element.geometry), animationStyle(element)]"
      :title="element.hyperlink?.tooltip || element.name"
      @click="onElementClick($event, element.id, element.hyperlink)"
      @dblclick="onElementDoubleClick($event, element.id)"
      @mouseenter="onElementMouseEnter($event, element.id)"
      @mouseleave="onElementMouseLeave($event, element.id)"
      @keydown.enter.space.stop.prevent="onElementKeyActivate($event, element.id, element.hyperlink)"
    >
      <svg v-if="element.kind === 'shape' && element.geometry === 'custom' && element.customPaths?.length" class="custom-geometry" viewBox="0 0 1000 1000" preserveAspectRatio="none" :style="customGeometryStyle(element)" aria-hidden="true">
        <defs>
          <template v-for="gradient in customGeometryGradients(element)" :key="gradient.id">
            <linearGradient v-if="gradient.kind === 'linear'" :id="gradient.id" gradientUnits="userSpaceOnUse" :x1="gradient.x1" :y1="gradient.y1" :x2="gradient.x2" :y2="gradient.y2"><stop v-for="(stop, stopIndex) in gradient.stops" :key="stopIndex" :offset="stop.offset" :stop-color="stop.color" :stop-opacity="stop.opacity" /></linearGradient>
            <radialGradient v-else :id="gradient.id" gradientUnits="userSpaceOnUse" :cx="gradient.cx" :cy="gradient.cy" :r="gradient.r"><stop v-for="(stop, stopIndex) in gradient.stops" :key="stopIndex" :offset="stop.offset" :stop-color="stop.color" :stop-opacity="stop.opacity" /></radialGradient>
          </template>
        </defs>
        <path v-for="(path, pathIndex) in element.customPaths" :key="pathIndex" :d="path.d" :fill="customGeometryFill(element, path)" :stroke="customGeometryStroke(element, path)" :stroke-width="customGeometryStrokeWidth(element)" :style="customGeometryStrokeStyle(element)" vector-effect="non-scaling-stroke" />
      </svg>
      <div v-if="element.kind === 'video' && element.mediaUrl && !thumbnail" class="image-viewport">
        <video class="slide-video" :data-media-id="element.id" :data-media-slide-count="mediaSlideCounts.get(element.id) || undefined" :src="element.mediaUrl" :poster="element.imageUrl" controls playsinline preload="metadata" :aria-label="element.name" :style="imageStyle(element)" @loadedmetadata="updateMediaPlayback($event, element)" @seeking="updateMediaPlayback($event, element)" @timeupdate="updateMediaPlayback($event, element)" @play="updateMediaOnPlay($event, element)" @pause="updateMediaFadeForEvent($event, element)" @ended="finishTimedMedia(element)" @click.stop @pointerdown.stop />
        <span v-if="invalidMediaTrimIds.includes(element.id)" class="media-warning" role="status">Trim range exceeds media duration; full media is used.</span>
        <span v-if="invalidMediaFadeIds.includes(element.id)" class="media-warning" role="status">Media fade duration exceeds the playable range; fades are ignored.</span>
        <span v-if="blockedMediaPlaybackIds.includes(element.id)" class="media-warning" role="status">The browser blocked automatic playback; use the media controls.</span>
      </div>
      <div v-else-if="element.kind === 'audio' && element.mediaUrl && !thumbnail" class="audio-viewport">
        <audio class="slide-audio" :data-media-id="element.id" :data-media-slide-count="mediaSlideCounts.get(element.id) || undefined" :src="element.mediaUrl" controls preload="metadata" :aria-label="element.name" @loadedmetadata="updateMediaPlayback($event, element)" @seeking="updateMediaPlayback($event, element)" @timeupdate="updateMediaPlayback($event, element)" @play="updateMediaOnPlay($event, element)" @pause="updateMediaFadeForEvent($event, element)" @volumechange="updateMediaFadeForEvent($event, element)" @ended="finishTimedMedia(element)" @click.stop @pointerdown.stop />
        <span v-if="invalidMediaTrimIds.includes(element.id)" class="media-warning" role="status">Trim range exceeds media duration; full media is used.</span>
        <span v-if="invalidMediaFadeIds.includes(element.id)" class="media-warning" role="status">Media fade duration exceeds the playable range; fades are ignored.</span>
        <span v-if="blockedMediaPlaybackIds.includes(element.id)" class="media-warning" role="status">The browser blocked automatic playback; use the media controls.</span>
      </div>
      <svg v-else-if="element.kind === 'chart' && element.chart" class="slide-chart" viewBox="0 0 1000 600" preserveAspectRatio="none" role="img" :aria-label="element.chart.title || element.name">
        <title>{{ element.chart.title || element.name }}</title>
        <defs>
          <clipPath :id="`chart-plot-${element.id}`"><polygon :points="`${chartPlot(element.id).left},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top + chartPlot(element.id).height} ${chartPlot(element.id).left},${chartPlot(element.id).top + chartPlot(element.id).height}`" /></clipPath>
          <clipPath v-if="chartVisuals.get(element.id)?.legendClip" :id="chartVisuals.get(element.id)?.legendClip?.id"><rect :x="chartVisuals.get(element.id)?.legendClip?.x" :y="chartVisuals.get(element.id)?.legendClip?.y" :width="chartVisuals.get(element.id)?.legendClip?.width" :height="chartVisuals.get(element.id)?.legendClip?.height" /></clipPath>
        </defs>
        <text v-if="element.chart.title" x="500" y="42" text-anchor="middle" class="chart-title">{{ element.chart.title }}</text>
        <g v-if="element.chart.type === 'pie' || element.chart.type === 'doughnut'">
          <path v-for="(slice, index) in chartVisuals.get(element.id)?.slices" :key="index" :d="slice.d" :fill="slice.color" :fill-rule="element.chart.type === 'doughnut' ? 'evenodd' : 'nonzero'" stroke="white" stroke-width="2" />
        </g>
        <g v-else>
          <template v-if="!element.chart.valueAxis?.deleted && element.chart.valueAxis?.majorGridlines"><line v-for="(tick, index) in chartVisuals.get(element.id)?.ticks" :key="`grid-${index}`" class="chart-major-gridline" :x1="element.chart.direction === 'horizontal' ? tick.x : chartPlot(element.id).left" :x2="element.chart.direction === 'horizontal' ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width" :y1="element.chart.direction === 'horizontal' ? chartPlot(element.id).top : tick.y" :y2="element.chart.direction === 'horizontal' ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y" stroke="#d1d5db" stroke-width="1" /></template>
          <template v-for="(tick, index) in chartVisuals.get(element.id)?.ticks" :key="`tick-${index}`"><text v-if="tick.showLabel" :x="tick.labelX" :y="tick.labelY" :text-anchor="tick.labelAnchor" class="chart-label chart-value-axis-label">{{ tick.label }}</text></template>
          <line v-for="(tick, index) in chartVisuals.get(element.id)?.axisTicks" :key="`axis-tick-${index}`" class="chart-axis-tick" :data-axis="tick.axis" :x1="tick.x1" :y1="tick.y1" :x2="tick.x2" :y2="tick.y2" stroke="#64748b" stroke-width="1.5" />
          <g :clip-path="`url(#chart-plot-${element.id})`">
            <line v-if="!element.chart.categoryAxisDeleted && chartVisuals.get(element.id)?.baseline" class="chart-category-axis-line" :x1="chartVisuals.get(element.id)?.baseline.x1" :y1="chartVisuals.get(element.id)?.baseline.y1" :x2="chartVisuals.get(element.id)?.baseline.x2" :y2="chartVisuals.get(element.id)?.baseline.y2" stroke="#64748b" stroke-width="2" />
            <line v-if="!element.chart.valueAxis?.deleted && chartVisuals.get(element.id)?.valueAxisLine" class="chart-value-axis-line" :x1="chartVisuals.get(element.id)?.valueAxisLine?.x1" :y1="chartVisuals.get(element.id)?.valueAxisLine?.y1" :x2="chartVisuals.get(element.id)?.valueAxisLine?.x2" :y2="chartVisuals.get(element.id)?.valueAxisLine?.y2" stroke="#64748b" stroke-width="2" />
            <rect v-for="(bar, index) in chartVisuals.get(element.id)?.bars" :key="`bar-${index}`" :x="bar.x" :y="bar.y" :width="bar.width" :height="bar.height" :fill="bar.color" />
            <path v-for="(line, index) in chartVisuals.get(element.id)?.lines" :key="`line-${index}`" :d="line.d" fill="none" :stroke="line.color" stroke-width="4" stroke-linecap="round" />
            <circle v-for="(marker, index) in chartVisuals.get(element.id)?.markers" :key="`marker-${index}`" :cx="marker.x" :cy="marker.y" r="5" :fill="marker.color" />
          </g>
          <text v-for="(label, index) in chartVisuals.get(element.id)?.categories" :key="`category-${index}`" :x="label.x" :y="label.y" :text-anchor="label.anchor" class="chart-label chart-category-label">{{ label.text }}</text>
        </g>
        <text v-for="(label, index) in chartVisuals.get(element.id)?.dataLabels" :key="`data-label-${index}`" :x="label.x" :y="label.y" :text-anchor="label.anchor" class="chart-data-label"><tspan v-for="(line, lineIndex) in label.lines" :key="lineIndex" :x="label.x" :y="label.y + (lineIndex - (label.lines.length - 1) / 2) * 22">{{ line }}</tspan></text>
        <g class="chart-legend-group" :clip-path="chartVisuals.get(element.id)?.legendClip ? `url(#${chartVisuals.get(element.id)?.legendClip?.id})` : undefined">
          <g v-for="(entry, index) in chartVisuals.get(element.id)?.legend" :key="`legend-${index}`">
            <rect :x="entry.x" :y="entry.y - 12" width="12" height="12" :fill="entry.color" />
            <text :x="entry.x + 18" :y="entry.y" class="chart-label chart-legend-label">{{ entry.text }}</text>
          </g>
        </g>
      </svg>
      <div v-else-if="(element.kind === 'picture' || element.kind === 'video') && element.imageUrl" class="image-viewport" aria-hidden="true">
        <img class="slide-image" :src="element.imageUrl" :alt="element.name" :style="imageStyle(element)">
      </div>
      <table v-else-if="element.kind === 'table' && element.table" class="slide-table">
        <colgroup>
          <col v-for="(column, columnIndex) in element.table.columns" :key="columnIndex" :style="{ width: `${column / Math.max(element.table.columns.reduce((sum, width) => sum + width, 0), 1) * 100}%` }">
        </colgroup>
        <tbody>
          <tr v-for="(row, rowIndex) in element.table.rows" :key="rowIndex" :style="{ height: `${row.height / Math.max(element.table.rows.reduce((sum, item) => sum + item.height, 0), 1) * 100}%` }">
            <template v-for="(cell, cellIndex) in row.cells" :key="cellIndex">
              <td v-if="!cell.hidden" :colspan="cell.colSpan" :rowspan="cell.rowSpan" :data-autofit="cell.textAutoFitDynamic ? 'normal' : undefined" :data-row-index="rowIndex" :data-cell-index="cellIndex" :style="tableCellStyle(cell)">
                <div class="table-cell-text" :style="tableCellTextStyle(cell, row.height)">
                  <div v-for="(entry, paragraphIndex) in paragraphEntries(cell.paragraphs)" :key="paragraphIndex" class="text-paragraph" :style="paragraphStyle(entry.paragraph, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), cell.textLineSpacingReduction, cell.textColumnsRightToLeft && (cell.textColumnCount ?? 1) > 1 ? 'ltr' : undefined)">
                    <img v-if="entry.paragraph.bullet?.kind === 'image' && entry.paragraph.bullet.url" class="paragraph-picture-bullet" :src="entry.paragraph.bullet.url" :style="paragraphBulletStyle(entry.paragraph)" alt="" aria-hidden="true">
                    <span v-else-if="entry.marker" class="paragraph-marker" :style="paragraphBulletStyle(entry.paragraph)" aria-hidden="true">{{ entry.marker }}</span>
                    <template v-for="(run, runIndex) in entry.paragraph.runs" :key="runIndex">
                      <component v-for="(segment, segmentIndex) in runSegments(run, element, paragraphIndex, runIndex)" :key="segmentIndex" :is="canActivateHyperlink(run.hyperlink) ? 'a' : 'span'" :href="canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : undefined" :title="run.hyperlink?.tooltip" :data-pptx-hyperlink="canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : undefined" :class="['text-run', { 'text-run-hyperlink': canActivateHyperlink(run.hyperlink) }]" :data-char-start="segment.start" :data-char-end="segment.end" :style="[runStyle(element, run, segment.script, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), paragraphIndex, segment.start, segment.end), textColorAnimationStyle(element, paragraphIndex, segment.start, segment.end), textRangeOpacityAnimationStyle(element, paragraphIndex, segment.start, segment.end), characterAnimationStyle(element, paragraphIndex, segment.start, segment.end), textRunAnimationStyle(element, paragraphIndex, segment.start, segment.end, run, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), segment.script), textTransformAnimationStyle(element, paragraphIndex, segment.start, segment.end)]">{{ segment.text }}</component>
                    </template>
                  </div>
                </div>
              </td>
            </template>
          </tr>
        </tbody>
      </table>
      <span v-else-if="element.kind === 'placeholder'">{{ element.mediaType ? 'Media unavailable' : 'Image unavailable' }}</span>
      <div v-if="element.paragraphs.length" class="text-frame" :data-element-id="element.id" :data-autofit="element.textAutoFit === 'shape' ? 'shape' : element.textAutoFitDynamic ? 'normal' : undefined" :style="textFrameStyle(element)">
        <div
          v-for="(entry, index) in paragraphEntries(element.paragraphs)"
          :key="index"
          class="text-paragraph"
          :style="[paragraphStyle(entry.paragraph, effectiveTextFontScale(element), element.textLineSpacingReduction, usesRightToLeftColumns(element) ? 'ltr' : undefined), paragraphAnimationStyle(element, index)]"
        >
          <img v-if="entry.paragraph.bullet?.kind === 'image' && entry.paragraph.bullet.url" class="paragraph-picture-bullet" :src="entry.paragraph.bullet.url" :style="paragraphBulletStyle(entry.paragraph)" alt="" aria-hidden="true">
          <span v-else-if="entry.marker" class="paragraph-marker" :style="paragraphBulletStyle(entry.paragraph)" aria-hidden="true">{{ entry.marker }}</span>
          <template v-for="(run, runIndex) in entry.paragraph.runs" :key="runIndex">
            <component v-for="(segment, segmentIndex) in runSegments(run, element, index, runIndex)" :key="segmentIndex" :is="canActivateHyperlink(run.hyperlink) ? 'a' : 'span'" :href="canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : undefined" :title="run.hyperlink?.tooltip" :data-pptx-hyperlink="canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : undefined" :class="['text-run', { 'text-run-hyperlink': canActivateHyperlink(run.hyperlink) }]" :data-char-start="segment.start" :data-char-end="segment.end" :style="[runStyle(element, run, segment.script, effectiveTextFontScale(element), index, segment.start, segment.end), textColorAnimationStyle(element, index, segment.start, segment.end), textRangeOpacityAnimationStyle(element, index, segment.start, segment.end), characterAnimationStyle(element, index, segment.start, segment.end), textRunAnimationStyle(element, index, segment.start, segment.end, run, effectiveTextFontScale(element), segment.script), textTransformAnimationStyle(element, index, segment.start, segment.end)]">{{ segment.text }}</component>
          </template>
        </div>
      </div>
      <svg v-if="element.kind === 'line'" class="line-render" :style="lineRenderStyle(element)" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true">
        <defs>
          <marker v-if="element.lineHeadEnd && element.lineHeadEnd.type !== 'none'" :id="lineEndId(element, 'head')" :markerWidth="lineEndMarkerSize(element, 'head').width" :markerHeight="lineEndMarkerSize(element, 'head').height" viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto-start-reverse">
            <path :d="lineEndPath(element.lineHeadEnd.type)" :fill="element.lineHeadEnd.type === 'arrow' ? 'none' : 'currentColor'" :stroke="element.lineHeadEnd.type === 'arrow' ? 'currentColor' : 'none'" />
          </marker>
          <marker v-if="element.lineTailEnd && element.lineTailEnd.type !== 'none'" :id="lineEndId(element, 'tail')" :markerWidth="lineEndMarkerSize(element, 'tail').width" :markerHeight="lineEndMarkerSize(element, 'tail').height" viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto">
            <path :d="lineEndPath(element.lineTailEnd.type)" :fill="element.lineTailEnd.type === 'arrow' ? 'none' : 'currentColor'" :stroke="element.lineTailEnd.type === 'arrow' ? 'currentColor' : 'none'" />
          </marker>
        </defs>
        <line x1="0" y1="0" :x2="element.width > 0 ? 1000 : 0" :y2="element.height > 0 ? 1000 : 0" :stroke="lineStrokeColor(element)" :stroke-width="Math.max(element.strokeWidth * scale, 1)" :style="customGeometryStrokeStyle(element)" :marker-start="lineEndUrl(element, 'head')" :marker-end="lineEndUrl(element, 'tail')" vector-effect="non-scaling-stroke" />
      </svg>
    </div>
    <span v-if="backgroundTileWarning" class="media-warning" role="status">{{ backgroundTileWarning }}</span>
  </div>
</template>

<style scoped>
.image-viewport { position: relative; width: 100%; height: 100%; overflow: hidden; }
.custom-geometry { position: absolute; inset: 0; width: 100%; height: 100%; overflow: visible; pointer-events: none; }
.slide-image { position: absolute; max-width: none; }
.slide-video { position: absolute; max-width: none; object-fit: fill; }
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
.paragraph-marker { display: inline-block; width: 1em; margin-left: -1em; margin-right: .45em; text-align: right; white-space: pre; }
.is-thumbnail .text-frame { overflow: hidden; }
@media (prefers-reduced-motion: reduce) { .slide-element { transition: none !important; } }
</style>
