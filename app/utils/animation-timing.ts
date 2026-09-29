export interface TimingCurve {
  acceleration?: number
  deceleration?: number
}

export function timingProgress(curve: TimingCurve, progress: number): number {
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

export function inverseTimingProgress(curve: TimingCurve, progress: number): number {
  if (!(curve.acceleration || curve.deceleration) || progress <= 0 || progress >= 1) return progress
  let low = 0
  let high = 1
  for (let step = 0; step < 28; step++) {
    const middle = (low + high) / 2
    if (timingProgress(curve, middle) < progress) low = middle
    else high = middle
  }
  return (low + high) / 2
}
