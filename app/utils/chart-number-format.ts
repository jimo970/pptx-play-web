interface NumberFormatOptions {
  useGrouping: boolean
  minimumFractionDigits: number
  maximumFractionDigits: number
  style?: 'decimal' | 'percent'
}

function numberFormatOptions(formatCode: string): NumberFormatOptions | undefined {
  const match = /^([#0]+(?:,[#0]{3})*)(?:\.([#0]+))?(%)?$/.exec(formatCode)
  if (!match) return undefined
  const fraction = match[2] || ''
  return {
    useGrouping: match[1]!.includes(','),
    minimumFractionDigits: [...fraction].filter(char => char === '0').length,
    maximumFractionDigits: fraction.length,
    ...(match[3] ? { style: 'percent' as const } : {}),
  }
}

export function isSupportedChartNumberFormat(formatCode: string): boolean {
  return formatCode.toLowerCase() === 'general' || numberFormatOptions(formatCode) !== undefined
}

export function formatChartNumber(value: number, formatCode?: string): string {
  if (!formatCode || formatCode.toLowerCase() === 'general') return String(value)
  const options = numberFormatOptions(formatCode)
  return options ? new Intl.NumberFormat(undefined, options).format(value) : String(value)
}
