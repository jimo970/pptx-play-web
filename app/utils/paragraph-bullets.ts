import type { PptxParagraph } from './pptx-model'

function alphabeticNumber(value: number, alphabet = 'abcdefghijklmnopqrstuvwxyz'): string {
  let number = Math.max(1, value)
  let result = ''
  while (number > 0) {
    number--
    result = alphabet.charAt(number % alphabet.length) + result
    number = Math.floor(number / alphabet.length)
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

function scriptDigits(value: number, digits: string): string {
  return String(value).replace(/\d/g, digit => digits.charAt(Number(digit)))
}

function chineseNumber(value: number, traditional: boolean): string {
  if (!Number.isSafeInteger(value) || value <= 0) return String(value)
  const digits = '零一二三四五六七八九'
  const units = traditional ? ['', '萬', '億', '兆', '京'] : ['', '万', '亿', '兆', '京']
  const groupUnits = ['', '十', '百', '千']
  const groups: number[] = []
  for (let remaining = value; remaining > 0; remaining = Math.floor(remaining / 10000)) groups.unshift(remaining % 10000)

  let result = ''
  let skippedGroup = false
  groups.forEach((group, index) => {
    if (!group) {
      if (result) skippedGroup = true
      return
    }
    if (result && (skippedGroup || group < 1000)) result += digits[0]
    const omitLeadingTen = !result && group >= 10 && group < 20
    let section = ''
    let pendingZero = false
    for (let position = 3; position >= 0; position--) {
      const digit = Math.floor(group / 10 ** position) % 10
      if (!digit) {
        if (section) pendingZero = true
        continue
      }
      if (pendingZero) section += digits[0]
      if (!(omitLeadingTen && position === 1 && digit === 1 && !section)) section += digits[digit]
      section += groupUnits[position]
      pendingZero = false
    }
    result += section + units[groups.length - index - 1]
    skippedGroup = false
  })
  return result
}

function numberedMarker(value: number, type: string): string {
  if (type === 'arabicPlain') return String(value)
  if (type === 'arabicDbPlain') return scriptDigits(value, '０１２３４５６７８９')
  let label = String(value)
  if (type.startsWith('ea1Chs')) label = chineseNumber(value, false)
  else if (type.startsWith('ea1Cht')) label = chineseNumber(value, true)
  else if (type.startsWith('thaiAlpha')) label = alphabeticNumber(value, 'กขคฆงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ')
  else if (type.startsWith('thaiNum')) label = scriptDigits(value, '๐๑๒๓๔๕๖๗๘๙')
  else if (type.startsWith('hindiNum')) label = scriptDigits(value, '०१२३४५६७८९')
  else if (type === 'arabicDbPeriod') return `${scriptDigits(value, '０１２３４５６７８９')}．`
  else if (type.startsWith('alpha')) {
    label = alphabeticNumber(value)
    if (type.startsWith('alphaUc')) label = label.toUpperCase()
  } else if (type.startsWith('roman')) {
    label = romanNumber(value)
    if (type.startsWith('romanUc')) label = label.toUpperCase()
  }
  if (type === 'ea1ChsPlain' || type === 'ea1ChtPlain') return label
  if (type.endsWith('ParenBoth')) return `(${label})`
  if (type.endsWith('ParenR')) return `${label})`
  return `${label}.`
}

// Numbering sequences depend on both the scheme and paragraph level.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.autonumberedbullet
export function paragraphEntries(paragraphs: PptxParagraph[]): { paragraph: PptxParagraph; marker: string }[] {
  const counters = new Map<string, number>()
  return paragraphs.map((paragraph) => {
    const bullet = paragraph.bullet
    if (!bullet || bullet.kind !== 'number') {
      counters.clear()
      return { paragraph, marker: bullet?.kind === 'char' ? bullet.char : '' }
    }
    const key = `${bullet.type}:${paragraph.level || 0}`
    const number = counters.get(key) ?? bullet.startAt ?? 1
    counters.set(key, number + 1)
    return { paragraph, marker: numberedMarker(number, bullet.type) }
  })
}
