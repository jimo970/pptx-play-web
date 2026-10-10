import type { PptxEmbeddedFontFace } from './pptx-model'
import { normalizedFontFamily } from './text-fonts'

export async function registerPptxEmbeddedFonts(embeddedFonts: PptxEmbeddedFontFace[]) {
  const warnings = new Set<string>()
  const faces: FontFace[] = []
  if (!embeddedFonts.length) return { faces, warnings: [] as string[] }
  if (typeof window === 'undefined' || typeof FontFace === 'undefined' || !window.document.fonts) {
    return { faces, warnings: ['This browser cannot load embedded fonts; system font fallbacks are in use.'] }
  }
  if (embeddedFonts.length > 256) warnings.add('Only the first 256 embedded font styles were loaded.')

  const seen = new Set<string>()
  await Promise.all(embeddedFonts.slice(0, 256).map(async embedded => {
    const family = normalizedFontFamily(embedded.family)
    if (!family) {
      warnings.add('An embedded font has an unusable typeface name and was skipped.')
      return
    }
    if (!['400', '700'].includes(embedded.weight) || !['normal', 'italic'].includes(embedded.style)) {
      warnings.add(`The embedded font "${family}" has an unsupported style descriptor and was skipped.`)
      return
    }
    try {
      const source = new URL(embedded.url)
      if (source.protocol !== 'blob:' || source.origin !== window.location.origin) {
        warnings.add(`The embedded font "${family}" does not reference a local font resource and was skipped.`)
        return
      }
    } catch {
      warnings.add(`The embedded font "${family}" has an invalid resource URL and was skipped.`)
      return
    }
    const key = `${family.toLowerCase()}\u0000${embedded.weight}\u0000${embedded.style}`
    if (seen.has(key)) {
      warnings.add(`Duplicate embedded font style for "${family}" was ignored.`)
      return
    }
    seen.add(key)
    try {
      const face = new FontFace(family, `url(${JSON.stringify(embedded.url)})`, {
        weight: embedded.weight,
        style: embedded.style,
      })
      await face.load()
      window.document.fonts.add(face)
      faces.push(face)
    } catch {
      warnings.add(`The embedded ${embedded.style} ${embedded.weight} font "${family}" could not be loaded; a system fallback is in use.`)
    }
  }))
  return { faces, warnings: [...warnings] }
}

export function releasePptxFontFaces(faces: FontFace[]) {
  if (typeof document === 'undefined') return
  for (const face of faces) document.fonts.delete(face)
}
