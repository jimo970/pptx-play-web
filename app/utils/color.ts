export function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  hue = ((hue % 1) + 1) % 1
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation
  const x = chroma * (1 - Math.abs((hue * 6) % 2 - 1))
  const match = lightness - chroma / 2
  const [r, g, b] = hue < 1 / 6 ? [chroma, x, 0] : hue < 2 / 6 ? [x, chroma, 0] : hue < 3 / 6 ? [0, chroma, x] : hue < 4 / 6 ? [0, x, chroma] : hue < 5 / 6 ? [x, 0, chroma] : [chroma, 0, x]
  return [(r + match) * 255, (g + match) * 255, (b + match) * 255]
}
