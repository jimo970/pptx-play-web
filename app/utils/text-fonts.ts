export type ScriptFont = 'latin' | 'eastAsia' | 'complexScript'

const eastAsian = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303f\uff00-\uffef]/u
const complex = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Thai}\p{Script=Lao}\p{Script=Tibetan}\p{Script=Myanmar}\p{Script=Khmer}\p{Script=Mongolian}]/u

export function characterScript(character: string, previous: ScriptFont = 'latin'): ScriptFont {
  return eastAsian.test(character) ? 'eastAsia' : complex.test(character) ? 'complexScript' : /\p{Mark}/u.test(character) ? previous : 'latin'
}

export function scriptSegments(text: string): { text: string; script: ScriptFont }[] {
  const segments: { text: string; script: ScriptFont }[] = []
  let previous: ScriptFont = 'latin'
  for (const character of text) {
    const script = characterScript(character, previous)
    const last = segments.at(-1)
    if (last?.script === script) last.text += character
    else segments.push({ text: character, script })
    previous = script
  }
  return segments
}

export function normalizedFontFamily(fontFamily: string | undefined): string {
  return (fontFamily || '').replace(/[\u0000-\u001f\u007f]/gu, ' ').trim().slice(0, 128)
}

export function cssFontFamily(fontFamily: string | undefined): string {
  const fontFamilyValue = normalizedFontFamily(fontFamily)
  const escapedFontFamily = fontFamilyValue.replace(/\\/gu, '\\\\').replace(/'/gu, "\\'")
  return escapedFontFamily ? `'${escapedFontFamily}', Arial, sans-serif` : 'Arial, sans-serif'
}
