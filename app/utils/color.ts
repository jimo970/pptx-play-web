export function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  hue = ((hue % 1) + 1) % 1
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation
  const x = chroma * (1 - Math.abs((hue * 6) % 2 - 1))
  const match = lightness - chroma / 2
  const [r, g, b] = hue < 1 / 6 ? [chroma, x, 0] : hue < 2 / 6 ? [x, chroma, 0] : hue < 3 / 6 ? [0, chroma, x] : hue < 4 / 6 ? [0, x, chroma] : hue < 5 / 6 ? [x, 0, chroma] : [chroma, 0, x]
  return [(r + match) * 255, (g + match) * 255, (b + match) * 255]
}

export function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const r = red / 255
  const g = green / 255
  const b = blue / 255
  const maximum = Math.max(r, g, b)
  const minimum = Math.min(r, g, b)
  const delta = maximum - minimum
  const lightness = (maximum + minimum) / 2
  if (delta === 0) return [0, 0, lightness]

  const saturation = delta / (1 - Math.abs(2 * lightness - 1))
  const hue = maximum === r
    ? ((g - b) / delta + (g < b ? 6 : 0)) / 6
    : maximum === g
      ? ((b - r) / delta + 2) / 6
      : ((r - g) / delta + 4) / 6
  return [hue, saturation, lightness]
}
