export type TextIterationGranularity = 'grapheme' | 'word'

export interface TextIterationRun {
    text: string
    language?: string
}

export interface TextIterationSegment {
    index: number
    segment: string
    isWordLike?: boolean
}

export function isTextIterationUnit(type: string | null, segment: {segment: string; isWordLike?: boolean}): boolean {
    if (type === 'wd') return segment.isWordLike === true
    if (type === 'lt') return !/^\s+$/u.test(segment.segment)
    return false
}

export function textIterationRanges(type: string | null, segments: Iterable<TextIterationSegment>): Array<{start: number; end: number}> {
    const ranges: Array<{start: number; end: number}> = []
    let lastWordRange: {start: number; end: number} | undefined
    let pendingPrefixStart: number | undefined

    for (const segment of segments) {
        const start = segment.index
        const end = start + segment.segment.length
        if (type === 'lt') {
            if (isTextIterationUnit(type, segment)) ranges.push({start, end})
            continue
        }
        if (type !== 'wd') continue
        if (/^\s+$/u.test(segment.segment)) {
            lastWordRange = undefined
            pendingPrefixStart = undefined
            continue
        }
        if (segment.isWordLike === true) {
            const range = {start: pendingPrefixStart ?? start, end}
            ranges.push(range)
            lastWordRange = range
            pendingPrefixStart = undefined
            continue
        }

        if (lastWordRange && lastWordRange.end === start) lastWordRange.end = end
        else pendingPrefixStart ??= start
    }

    return ranges
}

export function textIterationLocale(runs: readonly TextIterationRun[]): string | undefined {
    for (const run of runs) {
        const language = run.language?.trim()
        if (!language || !run.text.trim()) continue
        try {
            const locale = new Intl.Locale(language)
            if (locale.language) return locale.baseName
        } catch {
            // Invalid run tags are already reported while parsing the text style.
        }
    }
    return undefined
}

export function createTextIterationSegmenter(locale: string | undefined, granularity: TextIterationGranularity): Intl.Segmenter {
    try {
        return new Intl.Segmenter(locale, {granularity})
    } catch {
        return new Intl.Segmenter(undefined, {granularity})
    }
}
