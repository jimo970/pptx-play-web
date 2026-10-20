import type { Ref } from 'vue'
import type { MessageKey } from '../../i18n/locales/en'

export type UiLocale = 'en' | 'zh'

const errorKeys: Record<string, MessageKey> = {
  'Please choose a .pptx file. Legacy .ppt files are not supported.': 'choosePptxError',
  'This presentation is larger than the 150 MB browser limit.': 'fileTooLarge',
  'This file is encrypted or uses the legacy .ppt format. Save it as an unencrypted .pptx first.': 'encryptedOrLegacyPpt',
  'This presentation contains more than 20,000 package entries and exceeds the browser safety limit.': 'packageTooLarge',
  'This presentation has incomplete or invalid ZIP size metadata and cannot be read safely.': 'invalidZipSize',
  'This presentation expands beyond the 250 MB browser safety limit.': 'expandedFileTooLarge',
  'The PPTX file appears damaged or incomplete.': 'pptxDamaged',
  'Password-protected presentations are not supported.': 'passwordProtected',
  'The presentation could not be read.': 'presentationCouldNotBeRead',
  'ppt/presentation.xml contains invalid XML.': 'invalidPresentationXml',
  'This file is not a valid PowerPoint presentation.': 'notPowerPoint',
  'ppt/presentation.xml contains an invalid slide size.': 'invalidSlideSize',
  'No slides were found in this presentation.': 'noSlides',
  'This presentation contains more than 500 slides.': 'tooManySlides',
  'No readable slides were found in this presentation.': 'noReadableSlides',
}

export type LocaleMessageKey = MessageKey

export function useLocale() {
  const { locale: activeLocale, t: translate } = useI18n({ useScope: 'global' })
  const locale = activeLocale as unknown as Ref<UiLocale>
  const localePreference = useCookie<string | null>('deckline-locale', {
    default: () => null,
    maxAge: 60 * 60 * 24 * 365,
    path: '/',
    sameSite: 'lax',
  })
  const preferredLocale = useState<UiLocale>('deckline-locale', () => {
    if (localePreference.value === 'zh' || localePreference.value === 'en') return localePreference.value
    const acceptedLanguages = import.meta.server
      ? useRequestHeaders(['accept-language'])['accept-language'] || ''
      : typeof navigator === 'undefined' ? '' : navigator.language
    return acceptedLanguages.split(',')[0]?.trim().toLowerCase().startsWith('zh') ? 'zh' : 'en'
  })

  if (locale.value !== preferredLocale.value) locale.value = preferredLocale.value

  watch(locale, value => {
    localePreference.value = value
    preferredLocale.value = value
  }, { immediate: true })

  function t(key: MessageKey, values: Record<string, string | number> = {}) {
    return translate(key, values)
  }

  function translateError(message: string) {
    return t(errorKeys[message] || 'presentationCouldNotBeOpened')
  }

  function translateWarning(message: string) {
    const warning = message.toLowerCase()
    if (/\b(chart|workbook|worksheet|cell|series|axis|data point|number format)\b/.test(warning)) return t('warningChartSummary')
    if (/\b(font|typeface)\b/.test(warning)) return t('warningFontSummary')
    if (/\b(audio|video|media|sound|playback)\b/.test(warning)) return t('warningMediaSummary')
    if (/\b(animation|transition|timing|sequence|motion|trigger|keyframe)\b/.test(warning)) return t('warningAnimationSummary')
    if (/\b(hyperlink|relationship|unsafe url|link)\b/.test(warning)) return t('warningLinkSummary')
    if (/\b(text|paragraph|shape|picture|image|drawingml|geometry|fill|bullet|table|svg|glow|shadow|color)\b/.test(warning)) return t('warningSlideSummary')
    return t('warningGeneralSummary')
  }

  function toggleLocale() {
    locale.value = locale.value === 'en' ? 'zh' : 'en'
  }

  return { locale, t, translateError, translateWarning, toggleLocale }
}
