import type { PptxChartGrouping, PptxChartSeries, PptxChartStackSegment } from './pptx-model'

// Keep synchronous stage helpers separate from the full parser to avoid bundling it into the initial UI chunk.
export function getPptxChartStacking(series: PptxChartSeries[], grouping: Exclude<PptxChartGrouping, 'clustered'>) {
    const categoryCount = Math.max(0, ...series.map(item => item.values.length))
    const segments = series.map(item => Array<PptxChartStackSegment | undefined>(item.values.length))
    let min = 0
    let max = 0
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        const values = series.map(item => item.values[categoryIndex])
        const positiveTotal = values.reduce<number>((total, value) => total + (value !== null && value !== undefined && value > 0 ? value : 0), 0)
        const negativeTotal = values.reduce<number>((total, value) => total + (value !== null && value !== undefined && value < 0 ? -value : 0), 0)
        let positive = 0
        let negative = 0
        values.forEach((value, seriesIndex) => {
            if (value === null || value === undefined) return
            const segmentValue = grouping === 'percentStacked' ? value / (value < 0 ? negativeTotal || 1 : positiveTotal || 1) : value
            if (value < 0) {
                const start = negative
                negative += segmentValue
                segments[seriesIndex]![categoryIndex] = {start, end: negative}
            } else {
                const start = positive
                positive += segmentValue
                segments[seriesIndex]![categoryIndex] = {start, end: positive}
            }
        })
        min = Math.min(min, negative)
        max = Math.max(max, positive)
    }
    return {segments, min, max}
}

export function presetEllipsePoint(width: number, height: number, angle: number): [number, number] {
    const radians = angle * Math.PI / 10_800_000
    const ellipseAngle = Math.atan2(width * Math.sin(radians), height * Math.cos(radians))
    return [500 + 500 * Math.cos(ellipseAngle), 500 + 500 * Math.sin(ellipseAngle)]
}
