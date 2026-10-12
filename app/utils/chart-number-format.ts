interface NumberFormatOptions {
  useGrouping: boolean
  minimumIntegerDigits: number
  minimumFractionDigits: number
  maximumFractionDigits: number
  style?: 'decimal' | 'percent'
}

type NumberFormatCondition = { operator: '<' | '<=' | '=' | '>' | '>=' | '<>'; value: number }

type FractionNumberFormat = { whole?: string; numerator: string; denominator: string; fixedDenominator?: number; maximumDenominator: number; percent: boolean }

interface NumberFormatSection {
  options: NumberFormatOptions
  parentheses: boolean
  prefix: string
  suffix: string
  scale: number
  color?: string
  colorIndex?: number
  condition?: NumberFormatCondition
  scientific?: { symbol: string; sign: '+' | '-'; minimumDigits: number; integerDigits: number }
  fraction?: FractionNumberFormat
}

type DateFormatToken = { kind: 'literal'; value: string } | { kind: 'date'; code: 'd' | 'm' | 'y' | 'h' | 's'; length: number; elapsed?: boolean } | { kind: 'fraction'; digits: number } | { kind: 'ampm'; short: boolean }

type ExcelCalendarDate = { year: number; month: number; day: number; fraction: number; fictitiousLeapDay?: boolean }

export interface FormattedChartNumber {
  text: string
  color?: string
}

const numberFormatColors: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#008000',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
}

// [Color1]–[Color56] use the legacy indexed palette entries 8–63.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-3.0.1
const defaultIndexedColors = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

function indexedColor(index: number, indexedColors?: string[]): string | undefined {
  const color = indexedColors?.[index + 7]
  const hex = color?.trim().replace(/^#/, '').slice(-6)
  return hex && /^[\da-f]{6}$/i.test(hex) ? `#${hex.toLowerCase()}` : `#${defaultIndexedColors[index - 1]!.toLowerCase()}`
}

// Excel defines these eight named colors as leading section modifiers.
// https://support.microsoft.com/en-us/excel/number-format-codes-in-excel-for-mac
function numberFormatColor(code: string | undefined, indexedColors?: string[]): string | undefined {
  if (!code) return undefined
  const match = /^\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color(?:[1-9]|[1-4]\d|5[0-6]))\]/i.exec(code)
  if (!match) return undefined
  const colorIndex = /^Color(\d+)$/i.exec(match[1]!)
  return colorIndex ? indexedColor(Number(colorIndex[1]), indexedColors) : numberFormatColors[match[1]!.toLowerCase()]
}

function calendarDateFromSerial(serial: number, date1904: boolean): ExcelCalendarDate | undefined {
  const serialDay = Math.floor(serial)
  const fraction = serial - serialDay
  if (!date1904 && serialDay === 60) return { year: 1900, month: 1, day: 29, fraction, fictitiousLeapDay: true }
  const adjustedDay = !date1904 && serialDay > 60 ? serialDay - 1 : serialDay
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)
  const date = new Date(epoch + adjustedDay * 86_400_000)
  if (!Number.isFinite(date.getTime())) return undefined
  return { year: date.getUTCFullYear(), month: date.getUTCMonth(), day: date.getUTCDate(), fraction }
}

function serialFromCalendarDate(date: ExcelCalendarDate, date1904: boolean): number | undefined {
  if (date.fictitiousLeapDay && date.year === 1900 && date.month === 1 && date.day === 29 && !date1904) return 60 + date.fraction
  const utc = new Date(0)
  utc.setUTCFullYear(date.year, date.month, date.day)
  const time = utc.getTime()
  if (!Number.isFinite(time)) return undefined
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)
  const serial = (time - epoch) / 86_400_000 + date.fraction
  return !date1904 && time >= Date.UTC(1900, 2, 1) ? serial + 1 : serial
}

function daysInMonth(year: number, month: number): number {
  if (month === 1) return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28
  return [3, 5, 8, 10].includes(month) ? 30 : 31
}

function calendarDateAtInterval(start: ExcelCalendarDate, interval: number, unit: 'months' | 'years'): ExcelCalendarDate | undefined {
  const monthOffset = interval * (unit === 'years' ? 12 : 1)
  const monthIndex = start.year * 12 + start.month + monthOffset
  if (!Number.isSafeInteger(monthOffset) || !Number.isSafeInteger(monthIndex)) return undefined
  const year = Math.floor(monthIndex / 12)
  const month = ((monthIndex % 12) + 12) % 12
  return { year, month, day: Math.min(start.day, daysInMonth(year, month)), fraction: start.fraction }
}

export function getChartDateAxisMajorTicks(minimum: number, maximum: number, majorUnit: number, majorTimeUnit: 'days' | 'months' | 'years', date1904 = false): number[] | undefined {
  if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || maximum < minimum || !Number.isFinite(majorUnit) || majorUnit <= 0) return undefined
  if (majorTimeUnit !== 'days' && (!Number.isSafeInteger(majorUnit) || majorUnit * (majorTimeUnit === 'years' ? 12 : 1) > Number.MAX_SAFE_INTEGER)) return undefined
  const start = majorTimeUnit === 'days' ? undefined : calendarDateFromSerial(minimum, date1904)
  if (majorTimeUnit !== 'days' && !start) return undefined
  const ticks: number[] = []
  for (let index = 0; index <= 100; index++) {
    const calendarDate = start && majorTimeUnit !== 'days' ? calendarDateAtInterval(start, majorUnit * index, majorTimeUnit) : undefined
    const tick = index === 0 ? minimum : majorTimeUnit === 'days'
      ? minimum + majorUnit * index
      : calendarDate ? serialFromCalendarDate(calendarDate, date1904) : undefined
    if (tick === undefined || !Number.isFinite(tick)) return ticks.length ? ticks : undefined
    if (tick > maximum + 1e-9) return ticks.length ? ticks : undefined
    if (index && tick <= ticks[index - 1]!) return undefined
    ticks.push(tick)
  }
  const nextCalendarDate = start && majorTimeUnit !== 'days' ? calendarDateAtInterval(start, majorUnit * 101, majorTimeUnit) : undefined
  const nextTick = majorTimeUnit === 'days'
    ? minimum + majorUnit * 101
    : nextCalendarDate ? serialFromCalendarDate(nextCalendarDate, date1904) : undefined
  return nextTick !== undefined && Number.isFinite(nextTick) && nextTick <= maximum + 1e-9 ? undefined : ticks
}

function parseDateFormat(formatCode: string): DateFormatToken[] | undefined {
  if (!formatCode || formatCode.length > 128) return undefined
  const tokens: DateFormatToken[] = []
  for (let index = 0; index < formatCode.length;) {
    const remaining = formatCode.slice(index)
    const ampm = /^(AM\/PM|A\/P)/i.exec(remaining)
    if (ampm) {
      tokens.push({ kind: 'ampm', short: ampm[1]!.length === 3 })
      index += ampm[1]!.length
      continue
    }
    const char = formatCode[index]!
    if (char === '"') {
      const end = formatCode.indexOf('"', index + 1)
      if (end < 0) return undefined
      tokens.push({ kind: 'literal', value: formatCode.slice(index + 1, end) })
      index = end + 1
      continue
    }
    if (char === '\\') {
      if (index + 1 >= formatCode.length) return undefined
      tokens.push({ kind: 'literal', value: formatCode[index + 1]! })
      index += 2
      continue
    }
    if (char === '_') {
      if (index + 1 >= formatCode.length) return undefined
      tokens.push({ kind: 'literal', value: ' ' })
      index += 2
      continue
    }
    if (char === '[') {
      const end = formatCode.indexOf(']', index + 1)
      if (end < 0) return undefined
      const bracket = formatCode.slice(index + 1, end)
      const elapsedToken = /^([hms])\1?$/i.exec(bracket)
      if (elapsedToken) {
        tokens.push({kind: 'date', code: elapsedToken[1]!.toLowerCase() as 'h' | 'm' | 's', length: bracket.length, elapsed: true})
        index = end + 1
        continue
      }
      if (/^[hms]+$/i.test(bracket)) return undefined
      if (!/^\$(?:[^\]]*)-[0-9a-f]+$/i.test(bracket) && !/^[a-z]+$/i.test(bracket) && !/^Color(?:[1-9]|[1-4]\d|5[0-6])$/i.test(bracket)) return undefined
      index = end + 1
      continue
    }
    if (/[dmyhs]/i.test(char)) {
      const code = char.toLowerCase() as 'd' | 'm' | 'y' | 'h' | 's'
      let end = index + 1
      while (formatCode[end]?.toLowerCase() === code) end++
      const length = end - index
      if ((code === 'd' && length > 4) || (code === 'm' && length > 5) || (code === 'y' && length !== 2 && length !== 4) || ((code === 'h' || code === 's') && length > 2)) return undefined
      tokens.push({ kind: 'date', code, length })
      index = end
      continue
    }
    if (char === '.') {
      const previousToken = tokens.at(-1)
      if (previousToken?.kind === 'date' && previousToken.code === 's') {
        const fraction = /^\.0{1,6}(?!0)/.exec(remaining)
        if (!fraction) return undefined
        tokens.push({kind: 'literal', value: '.'})
        tokens.push({kind: 'fraction', digits: fraction[0].length - 1})
        index += fraction[0].length
        continue
      }
    }
    if (/[/\-.,: ]/.test(char)) {
      tokens.push({ kind: 'literal', value: char })
      index++
      continue
    }
    return undefined
  }

  const dateTokens = tokens.filter((token): token is Extract<DateFormatToken, {kind: 'date'}> => token.kind === 'date')
  if (!dateTokens.length) return undefined
  const hasClock = dateTokens.some(token => token.code === 'h' || token.code === 's') || tokens.some(token => token.kind === 'ampm')
  const formattedTokens = tokens.map((token, index) => {
    if (token.kind !== 'date' || token.code !== 'm' || token.elapsed) return token
    const previous = tokens.slice(0, index).reverse().find(item => item.kind !== 'literal' || !/^[: ]+$/.test(item.value))
    const next = tokens.slice(index + 1).find(item => item.kind !== 'literal' || !/^[: ]+$/.test(item.value))
    return hasClock && (previous?.kind === 'date' && previous.code === 'h' || next?.kind === 'date' && next.code === 's')
      ? {...token, length: -token.length}
      : token
  })
  return formattedTokens
}

function formatDateNumber(value: number, tokens: DateFormatToken[], date1904: boolean): string | undefined {
  if (!Number.isFinite(value) || value < 0 || value > 2_958_465) return undefined
  const fractionDigits = tokens.find((token): token is Extract<DateFormatToken, {kind: 'fraction'}> => token.kind === 'fraction')?.digits
  const precisionDigits = Math.max(fractionDigits ?? 3, 3)
  const unitsPerSecond = 10 ** precisionDigits
  const unitsPerDay = 86_400 * unitsPerSecond
  const displayStep = 10 ** (precisionDigits - (fractionDigits ?? 3))
  let serialDay = Math.floor(value)
  let timeUnits = Math.round((value - serialDay) * unitsPerDay / displayStep) * displayStep
  if (timeUnits >= unitsPerDay) {
    serialDay++
    timeUnits = 0
  }
  const fictitiousLeapDay = !date1904 && serialDay === 60
  const adjustedDay = !date1904 && serialDay > 60 ? serialDay - 1 : serialDay
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)
  const date = new Date(base + adjustedDay * 86_400_000 + Math.floor(timeUnits * 1_000 / unitsPerSecond))
  if (!Number.isFinite(date.getTime())) return undefined
  const year = date.getUTCFullYear()
  const month = fictitiousLeapDay ? 1 : date.getUTCMonth()
  const day = fictitiousLeapDay ? 29 : date.getUTCDate()
  const weekdayDate = fictitiousLeapDay ? new Date(Date.UTC(1900, 2, 1)) : date
  const hour = date.getUTCHours()
  const minute = date.getUTCMinutes()
  const second = date.getUTCSeconds()
  const fractionValue = Math.floor((timeUnits % unitsPerSecond) / displayStep)
  const totalSeconds = serialDay * 86_400 + timeUnits / unitsPerSecond
  const ampm = tokens.some(token => token.kind === 'ampm')
  const monthDate = fictitiousLeapDay ? new Date(Date.UTC(1900, 1, 1)) : date
  const local = (options: Intl.DateTimeFormatOptions, target = date) => new Intl.DateTimeFormat(undefined, {...options, timeZone: 'UTC'}).format(target)

  return tokens.map(token => {
    if (token.kind === 'literal') return token.value
    if (token.kind === 'ampm') return token.short ? (hour < 12 ? 'A' : 'P') : (hour < 12 ? 'AM' : 'PM')
    if (token.kind === 'fraction') return String(fractionValue).padStart(token.digits, '0')
    if (token.elapsed) {
      const unitSeconds = token.code === 'h' ? 3_600 : token.code === 'm' ? 60 : 1
      const duration = String(Math.floor(totalSeconds / unitSeconds))
      return token.length === 1 ? duration : duration.padStart(2, '0')
    }
    if (token.length < 0) {
      const digits = -token.length
      return digits === 1 ? String(minute) : String(minute).padStart(2, '0')
    }
    if (token.code === 'd') {
      if (token.length === 1) return String(day)
      if (token.length === 2) return String(day).padStart(2, '0')
      return local({weekday: token.length === 3 ? 'short' : 'long'}, weekdayDate)
    }
    if (token.code === 'm') {
      if (token.length === 1) return String(month + 1)
      if (token.length === 2) return String(month + 1).padStart(2, '0')
      if (token.length === 3) return local({month: 'short'}, monthDate)
      if (token.length === 4) return local({month: 'long'}, monthDate)
      return local({month: 'narrow'}, monthDate)
    }
    if (token.code === 'y') return token.length === 2 ? String(year % 100).padStart(2, '0') : String(year)
    if (token.code === 'h') {
      const value = ampm ? hour % 12 || 12 : hour
      return token.length === 1 ? String(value) : String(value).padStart(2, '0')
    }
    return token.length === 1 ? String(second) : String(second).padStart(2, '0')
  }).join('')
}

function splitNumberFormatSections(formatCode: string): string[] | undefined {
  const sections: string[] = []
  let start = 0
  let quoted = false
  for (let index = 0; index < formatCode.length; index++) {
    const char = formatCode[index]!
    if (char === '\\' || (!quoted && (char === '_' || char === '*'))) {
      index++
      continue
    }
    if (char === '"') quoted = !quoted
    else if (char === ';' && !quoted) {
      sections.push(formatCode.slice(start, index))
      start = index + 1
    }
  }
  if (quoted) return undefined
  sections.push(formatCode.slice(start))
  return sections
}

function parseNumberFormatLiteral(value: string): string | undefined {
  let literal = ''
  let quoted = false
  for (let index = 0; index < value.length; index++) {
    const char = value[index]!
    if (char === '\\') {
      if (index + 1 >= value.length) return undefined
      literal += value[++index]!
    } else if (char === '"') quoted = !quoted
    else if (quoted) literal += char
    else if (char === '_') {
      if (index + 1 >= value.length) return undefined
      literal += ' '
      index++
    } else if (/^[ $+():^'{}<>=\/!&~€£¥¢₹₽-]$/.test(char)) literal += char
    else return undefined
  }
  return quoted ? undefined : literal
}

function formatFractionMaskInteger(value: number, mask: string): string {
  const minimumIntegerDigits = Math.max(1, [...mask].filter(char => char === '0').length)
  const formatted = new Intl.NumberFormat(undefined, {
    useGrouping: mask.includes(','),
    minimumIntegerDigits,
    maximumFractionDigits: 0,
  }).format(value)
  return mask.includes('?') ? formatted.padStart(mask.length, '\u00a0') : formatted
}

function greatestCommonDivisor(left: number, right: number): number {
  while (right) [left, right] = [right, left % right]
  return left
}

function formatFractionNumber(value: number, fraction: FractionNumberFormat): string | undefined {
  if (!Number.isFinite(value)) return undefined
  const magnitude = Math.abs(value) * (fraction.percent ? 100 : 1)
  const hasWhole = Boolean(fraction.whole)
  let whole = hasWhole ? Math.floor(magnitude) : 0
  if (!Number.isSafeInteger(whole)) return undefined
  const target = hasWhole ? magnitude - whole : magnitude
  let numerator = 0
  let denominator = fraction.fixedDenominator || 1

  if (fraction.fixedDenominator) {
    if (target > Number.MAX_SAFE_INTEGER / denominator) return undefined
    numerator = Math.round(target * denominator)
  } else {
    if (target > Number.MAX_SAFE_INTEGER / fraction.maximumDenominator) return undefined
    let bestError = Number.POSITIVE_INFINITY
    for (let candidate = 1; candidate <= fraction.maximumDenominator; candidate++) {
      const candidateNumerator = Math.round(target * candidate)
      const error = Math.abs(target - candidateNumerator / candidate)
      if (error < bestError - Number.EPSILON * Math.max(1, target)) {
        bestError = error
        numerator = candidateNumerator
        denominator = candidate
      }
    }
    if (numerator) {
      const divisor = greatestCommonDivisor(numerator, denominator)
      numerator /= divisor
      denominator /= divisor
    }
  }

  if (hasWhole && numerator >= denominator) {
    whole++
    numerator = 0
  }
  const wholeMask = fraction.whole
  const showWhole = Boolean(wholeMask && (whole > 0 || wholeMask.includes('0')))
  const wholeText = showWhole ? formatFractionMaskInteger(whole, wholeMask!) : ''
  const sign = value < 0 ? '-' : ''
  if (!numerator) return `${sign}${wholeText || '0'}`

  const numeratorText = formatFractionMaskInteger(numerator, fraction.numerator)
  const denominatorText = fraction.fixedDenominator
    ? new Intl.NumberFormat(undefined, {maximumFractionDigits: 0}).format(denominator)
    : formatFractionMaskInteger(denominator, fraction.denominator)
  const fractionText = `${numeratorText}/${denominatorText}`
  return `${sign}${wholeText ? `${wholeText} ` : ''}${fractionText}`
}

function parseNumberFormatSection(code: string, index: number): NumberFormatSection | null | undefined {
  if (!code) return null
  const colorMatch = /^\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color(?:[1-9]|[1-4]\d|5[0-6]))\]/i.exec(code)
  const colorIndex = colorMatch ? /^Color(\d+)$/i.exec(colorMatch[1]!) : undefined
  const color = colorMatch && !colorIndex ? numberFormatColors[colorMatch[1]!.toLowerCase()] : undefined
  if (colorMatch) code = code.slice(colorMatch[0].length)
  const conditionMatch = /^\[(<=|>=|<>|=|<|>)(-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)\]/.exec(code)
  const conditionValue = conditionMatch ? Number(conditionMatch[2]) : undefined
  if (conditionMatch && !Number.isFinite(conditionValue)) return undefined
  const condition = conditionMatch
    ? {operator: conditionMatch[1] as NumberFormatCondition['operator'], value: conditionValue!}
    : undefined
  if (conditionMatch) code = code.slice(conditionMatch[0].length)
  const parentheses = code.startsWith('(') && code.endsWith(')')
  if (parentheses && index !== 1) return undefined
  if (parentheses) code = code.slice(1, -1)

  let start = -1
  let quoted = false
  for (let offset = 0; offset < code.length; offset++) {
    const char = code[offset]!
    if (char === '\\' || (!quoted && char === '_')) {
      offset++
      continue
    }
    if (char === '"') quoted = !quoted
    else if (!quoted && (char === '#' || char === '0' || char === '?')) {
      start = offset
      break
    }
  }
  if (start < 0) return undefined
  const fractionMatch = /^(?:([#0?]+(?:,[#0?]{3})*) +)?([#0?]+)\/([#0?]+|[1-9]\d*)/.exec(code.slice(start))
  if (fractionMatch) {
    const whole = fractionMatch[1] || undefined
    const numerator = fractionMatch[2]!
    const denominator = fractionMatch[3]!
    const fixedDenominator = /^\d+$/.test(denominator) ? Number(denominator) : undefined
    const denominatorEnd = start + fractionMatch[0].length
    const rawSuffix = code.slice(denominatorEnd)
    const percentMatch = /^(\s*)%/.exec(rawSuffix)
    const percent = Boolean(percentMatch)
    const suffixStart = denominatorEnd + (percentMatch?.[0].length || 0)
    const prefix = parseNumberFormatLiteral(code.slice(0, start))
    const percentSpacing = percentMatch ? parseNumberFormatLiteral(percentMatch[1]!) : ''
    const remainingSuffix = parseNumberFormatLiteral(code.slice(suffixStart))
    if (prefix === undefined || percentSpacing === undefined || remainingSuffix === undefined || (!percent && rawSuffix.includes('%')) || numerator.length > 3 || (fixedDenominator === undefined && denominator.length > 3)
      || (whole && whole.replaceAll(',', '').length > 15)
      || (fixedDenominator !== undefined && (!Number.isSafeInteger(fixedDenominator) || fixedDenominator < 1 || fixedDenominator > 999_999))) return undefined
    const suffix = `${percentSpacing}${percent ? '%' : ''}${remainingSuffix}`
    return {
      options: {useGrouping: Boolean(whole?.includes(',')), minimumIntegerDigits: 1, minimumFractionDigits: 0, maximumFractionDigits: 0},
      parentheses,
      prefix,
      suffix,
      scale: 0,
      ...(color ? {color} : {}),
      ...(colorIndex ? {colorIndex: Number(colorIndex[1])} : {}),
      ...(condition ? {condition} : {}),
      fraction: {whole, numerator, denominator, ...(fixedDenominator !== undefined ? {fixedDenominator} : {}), maximumDenominator: fixedDenominator === undefined ? 10 ** denominator.length - 1 : 0, percent},
    }
  }
  let end = start
  while (/[#0,.%]/.test(code[end] || '')) end++
  const maskEnd = end
  let scientific: NumberFormatSection['scientific']
  if (/[Ee]/.test(code[end] || '')) {
    const exponent = /^([Ee])([+-])(0+)/.exec(code.slice(end))
    if (!exponent) return undefined
    const integerDigits = code.slice(start, end).split('.')[0]!.replaceAll(',', '').replaceAll('%', '').length
    scientific = {symbol: exponent[1]!, sign: exponent[2] as '+' | '-', minimumDigits: exponent[3]!.length, integerDigits}
    end += exponent[0].length
  }
  let mask = code.slice(start, maskEnd)
  const trailingPercent = Boolean(scientific && code[end] === '%')
  if (trailingPercent) end++
  const percentCount = [...mask].filter(char => char === '%').length + Number(trailingPercent)
  if (percentCount > 1) return undefined
  mask = mask.replace('%', '')
  let scale = 0
  while (mask.endsWith(',')) {
    scale++
    mask = mask.slice(0, -1)
  }
  if (scientific && scale) return undefined
  const match = /^([#0]+(?:,[#0]{3})*)(?:\.([#0]+))?$/.exec(mask)
  if (!match) return undefined
  if ((match[2]?.length || 0) > 100) return undefined
  const prefix = parseNumberFormatLiteral(code.slice(0, start))
  const suffix = parseNumberFormatLiteral(code.slice(end))
  if (prefix === undefined || suffix === undefined) return undefined
  const fraction = match[2] || ''
  return {
    options: {
      useGrouping: match[1]!.includes(','),
      minimumIntegerDigits: Math.max(1, [...match[1]!].filter(char => char === '0').length),
      minimumFractionDigits: [...fraction].filter(char => char === '0').length,
      maximumFractionDigits: fraction.length,
      ...(percentCount ? { style: 'percent' as const } : {}),
    },
    parentheses,
    prefix,
    suffix,
    scale,
    ...(color ? {color} : {}),
    ...(colorIndex ? {colorIndex: Number(colorIndex[1])} : {}),
    ...(condition ? {condition} : {}),
    ...(scientific ? {scientific} : {}),
  }
}

function conditionsCoverAllValues(left: NumberFormatCondition, right: NumberFormatCondition): boolean {
  if (left.value !== right.value) return false
  const operators = [left.operator, right.operator]
  return operators.includes('<=') && operators.includes('>')
    || operators.includes('<') && operators.includes('>=')
    || operators.includes('<>') && operators.includes('=')
}

function parseNumberFormat(formatCode: string): (NumberFormatSection | null)[] | undefined {
  const codes = splitNumberFormatSections(formatCode)
  if (!codes || codes.length > 3 || !codes[0]) return undefined

  const sections = codes.map(parseNumberFormatSection)

  if (sections.some(section => section === undefined)) return undefined
  const parsedSections = sections as (NumberFormatSection | null)[]
  const conditionalIndexes = parsedSections.flatMap((section, index) => section?.condition ? [index] : [])
  if (!conditionalIndexes.length) return parsedSections
  // No-fallback condition pairs are limited to complementary comparisons; other formats need an unconditional fallback.
  // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/17d11129-219b-4e2c-88db-45844d21e528
  const supportedConditions = codes.length === 2
    ? conditionalIndexes[0] === 0 && (conditionalIndexes.length === 1
      || conditionalIndexes[1] === 1 && parsedSections[0]?.condition !== undefined && parsedSections[1]?.condition !== undefined
        && conditionsCoverAllValues(parsedSections[0].condition, parsedSections[1].condition))
    : codes.length === 3
      ? conditionalIndexes.length === 2 && conditionalIndexes[0] === 0 && conditionalIndexes[1] === 1
      : false
  return supportedConditions && !parsedSections.some(section => section?.parentheses) ? parsedSections : undefined
}

function matchesNumberCondition(value: number, condition: NumberFormatCondition): boolean {
  switch (condition.operator) {
    case '<': return value < condition.value
    case '<=': return value <= condition.value
    case '=': return value === condition.value
    case '>': return value > condition.value
    case '>=': return value >= condition.value
    case '<>': return value !== condition.value
  }
}

export function isSupportedChartNumberFormat(formatCode: string): boolean {
  return formatCode.toLowerCase() === 'general' || parseNumberFormat(formatCode) !== undefined
    || splitNumberFormatSections(formatCode)?.some(section => parseDateFormat(section) !== undefined) === true
}

function formatScientificNumber(value: number, section: NumberFormatSection): string {
  const scientific = section.scientific!
  const integerDigits = scientific.integerDigits
  if (!Number.isFinite(value)) return String(value)
  let exponent = 0
  let mantissa = 0
  if (value !== 0) {
    const [coefficient, rawExponentText] = Math.abs(value).toExponential().split('e')
    const rawExponent = Number(rawExponentText) + (section.options.style === 'percent' ? 2 : 0)
    exponent = Math.floor(rawExponent / integerDigits) * integerDigits
    mantissa = Number(coefficient) * 10 ** (rawExponent - exponent)
  }
  const roundingScale = 10 ** section.options.maximumFractionDigits
  const roundedMantissa = Math.round(mantissa * roundingScale) / roundingScale
  if (roundedMantissa >= 10 ** integerDigits) {
    exponent += integerDigits
    mantissa /= 10 ** integerDigits
  }
  const formatted = new Intl.NumberFormat(undefined, {...section.options, style: 'decimal'}).format(Math.sign(value) * mantissa)
  const exponentSign = exponent < 0 ? '-' : scientific.sign === '+' ? '+' : ''
  const exponentText = `${scientific.symbol}${exponentSign}${String(Math.abs(exponent)).padStart(scientific.minimumDigits, '0')}`
  const percent = section.options.style === 'percent'
    ? new Intl.NumberFormat(undefined, {style: 'percent', maximumFractionDigits: 0}).formatToParts(0.01).find(part => part.type === 'percentSign')?.value || '%'
    : ''
  return `${formatted}${exponentText}${percent}`
}

export function formatChartNumberWithColor(value: number, formatCode?: string, date1904 = false, indexedColors?: string[]): FormattedChartNumber {
  if (!formatCode || formatCode.toLowerCase() === 'general') return {text: String(value)}
  const sections = splitNumberFormatSections(formatCode)
  const dateSection = sections && (value < 0 ? sections[1] : value === 0 ? sections[2] || sections[0] : sections[0])
  const dateTokens = dateSection ? parseDateFormat(dateSection) : undefined
  if (dateTokens) {
    const text = formatDateNumber(Math.abs(value), dateTokens, date1904) ?? String(value)
    const color = numberFormatColor(dateSection, indexedColors)
    return {text, ...(color ? {color} : {})}
  }
  const numberSections = parseNumberFormat(formatCode)
  if (!numberSections) return {text: String(value)}

  let section = numberSections[0]!
  let formattedValue = value
  if (numberSections.some(candidate => candidate?.condition)) {
    const matching = numberSections.slice(0, 2).find(candidate => candidate?.condition && matchesNumberCondition(value, candidate.condition))
    const fallback = numberSections.find(candidate => candidate && !candidate.condition)
    const selected = matching || fallback
    if (!selected) return {text: String(value)}
    section = selected
  } else {
    if (value < 0 && numberSections.length > 1) {
      const negative = numberSections[1]
      if (!negative) return {text: ''}
      section = negative
      formattedValue = Math.abs(value)
    } else if (value === 0 && numberSections.length > 2) {
      const zero = numberSections[2]
      if (!zero) return {text: ''}
      section = zero
    }
  }

  let scaledValue = formattedValue
  for (let index = 0; index < section.scale; index++) scaledValue /= 1000
  const formatted = section.fraction
    ? formatFractionNumber(scaledValue, section.fraction)
    : section.scientific
    ? formatScientificNumber(scaledValue, section)
    : new Intl.NumberFormat(undefined, section.options).format(scaledValue)
  if (formatted === undefined) return {text: String(value)}
  const valueText = `${section.prefix}${formatted}${section.suffix}`
  const text = section.parentheses ? `(${valueText})` : valueText
  const color = section.color ?? (section.colorIndex ? indexedColor(section.colorIndex, indexedColors) : undefined)
  return {text, ...(color ? {color} : {})}
}

export function formatChartNumber(value: number, formatCode?: string, date1904 = false, indexedColors?: string[]): string {
  return formatChartNumberWithColor(value, formatCode, date1904, indexedColors).text
}
