export interface TimingCurve {
  acceleration?: number
  deceleration?: number
  timeFilter?: Array<{input: number; output: number}>
  bounceEndRatio?: number
}

export function parseOpenXmlPercentage(raw: string | null | undefined): number | undefined {
  if (raw === null || raw === undefined) return undefined
  const value = raw.trim()
  // PresentationML ST_Percentage also uses percent-plus-integer compatibility forms, such as 56.7%457200.
  const percent = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))%(?:\d+)?$/.exec(value)
  const parsed = percent
    ? Number(percent[1]) / 100
    : /^[+-]?\d+$/.test(value) ? Number(value) / 100_000 : Number.NaN
  return Number.isFinite(parsed) ? parsed : undefined
}

export function parseAnimationSpeed(raw: string | null | undefined): number | undefined {
  if (raw === null || raw === undefined) return 1
  return parseOpenXmlPercentage(raw)
}

export function parseTimingFilter(raw: string | null | undefined): TimingCurve['timeFilter'] {
  if (!raw?.trim() || raw.length > 16_384) return undefined
  const pairs = raw.trim().split(';')
  if (pairs.length < 2 || pairs.length > 256) return undefined
  const points: NonNullable<TimingCurve['timeFilter']> = []
  for (const pair of pairs) {
    const values = pair.split(',')
    if (values.length !== 2 || values.some(value => !value.trim())) return undefined
    const input = Number(values[0])
    const output = Number(values[1])
    const previous = points.at(-1)
    if (!Number.isFinite(input) || !Number.isFinite(output) || input < 0 || input > 1 ||
      previous && input <= previous.input) return undefined
    points.push({input, output})
  }
  // MS-OI29500 constrains t to [0, 1] and describes t' as warped local time without a range bound.
  // Keep finite t' overshoot; filterProgress holds the nearest authored value outside the t domain.
  return points
}

function accelerationProgress(curve: TimingCurve, progress: number): number {
  const acceleration = curve.acceleration || 0
  const deceleration = curve.deceleration || 0
  if (!acceleration && !deceleration) return progress
  const area = 1 - (acceleration + deceleration) / 2
  if (acceleration && progress < acceleration) return progress ** 2 / (2 * acceleration * area)
  if (deceleration && progress > 1 - deceleration) {
    const tail = progress - (1 - deceleration)
    return (1 - deceleration - acceleration / 2 + tail - tail ** 2 / (2 * deceleration)) / area
  }
  return (acceleration / 2 + progress - acceleration) / area
}

function filterProgress(filter: NonNullable<TimingCurve['timeFilter']>, mapped: number): number {
  if (filter.length < 2) return mapped
  if (mapped <= filter[0]!.input) return filter[0]!.output
  if (mapped >= filter.at(-1)!.input) return filter.at(-1)!.output
  let low = 0
  let high = filter.length - 1
  while (high - low > 1) {
    const middle = (low + high) >>> 1
    if (filter[middle]!.input < mapped) low = middle
    else high = middle
  }
  const before = filter[low]!
  const after = filter[high]!
  return before.output + (after.output - before.output) * (mapped - before.input) / (after.input - before.input)
}

function inverseFilterProgress(
  filter: NonNullable<TimingCurve['timeFilter']>,
  progress: number,
  reverse = false,
  reverseDirection: 'up' | 'down' = 'down'
): number {
  if (filter.length < 2) return progress
  const initialProgress = filterProgress(filter, reverse ? 1 : 0)
  const alreadyReachedFromStart = reverseDirection === 'up'
    ? initialProgress >= progress
    : initialProgress <= progress
  if (alreadyReachedFromStart) return 0

  const authoredDirection = reverse
    ? reverseDirection === 'up' ? 'down' : 'up'
    : reverseDirection
  for (let step = 0; step < filter.length - 1; step++) {
    const index = reverse ? filter.length - 2 - step : step
    const before = filter[index]!
    const after = filter[index + 1]!
    // Traverse control points in wall-clock order. Include threshold equality at either end,
    // but require actual movement so a flat threshold segment is not treated as a crossing.
    const crosses = authoredDirection === 'up'
      ? before.output < after.output && before.output <= progress && after.output >= progress
      : before.output > after.output && before.output >= progress && after.output <= progress
    if (crosses) {
      const fraction = (progress - before.output) / (after.output - before.output)
      return before.input + (after.input - before.input) * fraction
    }
  }
  // The parent never reaches this child offset in the selected playback direction.
  return Number.POSITIVE_INFINITY
}

function inverseMonotonicTimingProgress(curve: TimingCurve, progress: number): number {
  if (progress <= 0 || progress >= 1) return progress
  let low = 0
  let high = 1
  for (let step = 0; step < 48; step++) {
    const middle = (low + high) / 2
    if (accelerationProgress(curve, middle) < progress) low = middle
    else high = middle
  }
  return (low + high) / 2
}

function baseTimingProgress(curve: TimingCurve, progress: number): number {
  const mapped = accelerationProgress(curve, progress)
  const filter = curve.timeFilter
  return filter ? filterProgress(filter, mapped) : mapped
}

export function timingProgress(curve: TimingCurve, progress: number): number {
  const bounceEndRatio = curve.bounceEndRatio
  if (bounceEndRatio === undefined || bounceEndRatio <= 0 || bounceEndRatio >= 1) {
    return baseTimingProgress(curve, progress)
  }

  const bounceStart = 1 - bounceEndRatio
  if (progress <= bounceStart) return baseTimingProgress(curve, progress / bounceStart)
  const bounceProgress = Math.max(0, Math.min(1, (progress - bounceStart) / bounceEndRatio))
  // MS-PPTX defines this as a mass-spring-damper bounce, but does not publish
  // the Office model parameters. Use a damped two-cycle approximation and keep
  // the final value exact so held animations settle on their requested endpoint.
  return 1 + 0.15 * Math.exp(-5 * bounceProgress) * Math.sin(4 * Math.PI * bounceProgress)
}

export function inverseTimingProgress(curve: TimingCurve, progress: number, reverse = false, reverseDirection: 'up' | 'down' = 'down'): number {
  const filter = curve.timeFilter
  if (!filter) return inverseMonotonicTimingProgress(curve, progress)
  const mappedProgress = inverseFilterProgress(filter, progress, reverse, reverseDirection)
  return inverseMonotonicTimingProgress({...curve, timeFilter: undefined}, mappedProgress)
}
