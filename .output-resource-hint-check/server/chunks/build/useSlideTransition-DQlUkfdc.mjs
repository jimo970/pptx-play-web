import { _ as _plugin_vue_export_helper_default } from '../virtual/entry.mjs';
import { watch, nextTick, defineComponent, useId, ref, computed, reactive, mergeProps, unref, createVNode, resolveDynamicComponent, withCtx, openBlock, createBlock, toDisplayString, Fragment, createTextVNode, useSSRContext } from 'vue';
import { ssrRenderAttrs, ssrRenderList, ssrRenderAttr, ssrRenderClass, ssrRenderStyle, ssrInterpolate, ssrRenderVNode } from 'vue/server-renderer';

function calendarDateFromSerial(serial, date1904) {
  const serialDay = Math.floor(serial);
  const fraction = serial - serialDay;
  if (!date1904 && serialDay === 60) return {
    year: 1900,
    month: 1,
    day: 29,
    fraction,
    fictitiousLeapDay: true
  };
  const adjustedDay = !date1904 && serialDay > 60 ? serialDay - 1 : serialDay;
  const date = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)) + adjustedDay * 864e5);
  if (!Number.isFinite(date.getTime())) return void 0;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth(),
    day: date.getUTCDate(),
    fraction
  };
}
function serialFromCalendarDate(date, date1904) {
  if (date.fictitiousLeapDay && date.year === 1900 && date.month === 1 && date.day === 29 && !date1904) return 60 + date.fraction;
  const utc = /* @__PURE__ */ new Date(0);
  utc.setUTCFullYear(date.year, date.month, date.day);
  const time = utc.getTime();
  if (!Number.isFinite(time)) return void 0;
  const serial = (time - (date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31))) / 864e5 + date.fraction;
  return !date1904 && time >= Date.UTC(1900, 2, 1) ? serial + 1 : serial;
}
function daysInMonth(year, month) {
  if (month === 1) return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28;
  return [
    3,
    5,
    8,
    10
  ].includes(month) ? 30 : 31;
}
function calendarDateAtInterval(start, interval, unit) {
  const monthOffset = interval * (unit === "years" ? 12 : 1);
  const monthIndex = start.year * 12 + start.month + monthOffset;
  if (!Number.isSafeInteger(monthOffset) || !Number.isSafeInteger(monthIndex)) return void 0;
  const year = Math.floor(monthIndex / 12);
  const month = (monthIndex % 12 + 12) % 12;
  return {
    year,
    month,
    day: Math.min(start.day, daysInMonth(year, month)),
    fraction: start.fraction
  };
}
function getChartDateAxisMajorTicks(minimum, maximum, majorUnit, majorTimeUnit, date1904 = false) {
  if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || maximum < minimum || !Number.isFinite(majorUnit) || majorUnit <= 0) return void 0;
  if (majorTimeUnit !== "days" && (!Number.isSafeInteger(majorUnit) || majorUnit * (majorTimeUnit === "years" ? 12 : 1) > Number.MAX_SAFE_INTEGER)) return void 0;
  const start = majorTimeUnit === "days" ? void 0 : calendarDateFromSerial(minimum, date1904);
  if (majorTimeUnit !== "days" && !start) return void 0;
  const ticks = [];
  for (let index = 0; index <= 100; index++) {
    const calendarDate = start && majorTimeUnit !== "days" ? calendarDateAtInterval(start, majorUnit * index, majorTimeUnit) : void 0;
    const tick = index === 0 ? minimum : majorTimeUnit === "days" ? minimum + majorUnit * index : calendarDate ? serialFromCalendarDate(calendarDate, date1904) : void 0;
    if (tick === void 0 || !Number.isFinite(tick)) return ticks.length ? ticks : void 0;
    if (tick > maximum + 1e-9) return ticks.length ? ticks : void 0;
    if (index && tick <= ticks[index - 1]) return void 0;
    ticks.push(tick);
  }
  const nextCalendarDate = start && majorTimeUnit !== "days" ? calendarDateAtInterval(start, majorUnit * 101, majorTimeUnit) : void 0;
  const nextTick2 = majorTimeUnit === "days" ? minimum + majorUnit * 101 : nextCalendarDate ? serialFromCalendarDate(nextCalendarDate, date1904) : void 0;
  return nextTick2 !== void 0 && Number.isFinite(nextTick2) && nextTick2 <= maximum + 1e-9 ? void 0 : ticks;
}
function parseDateFormat(formatCode) {
  var _a;
  if (!formatCode || formatCode.length > 128) return void 0;
  const tokens = [];
  for (let index = 0; index < formatCode.length; ) {
    const remaining = formatCode.slice(index);
    const ampm = /^(AM\/PM|A\/P)/i.exec(remaining);
    if (ampm) {
      tokens.push({
        kind: "ampm",
        short: ampm[1].length === 3
      });
      index += ampm[1].length;
      continue;
    }
    const char = formatCode[index];
    if (char === '"') {
      const end = formatCode.indexOf('"', index + 1);
      if (end < 0) return void 0;
      tokens.push({
        kind: "literal",
        value: formatCode.slice(index + 1, end)
      });
      index = end + 1;
      continue;
    }
    if (char === "\\") {
      if (index + 1 >= formatCode.length) return void 0;
      tokens.push({
        kind: "literal",
        value: formatCode[index + 1]
      });
      index += 2;
      continue;
    }
    if (char === "_") {
      if (index + 1 >= formatCode.length) return void 0;
      tokens.push({
        kind: "literal",
        value: " "
      });
      index += 2;
      continue;
    }
    if (char === "[") {
      const end = formatCode.indexOf("]", index + 1);
      if (end < 0) return void 0;
      const bracket = formatCode.slice(index + 1, end);
      const elapsedToken = /^([hms])\1?$/i.exec(bracket);
      if (elapsedToken) {
        tokens.push({
          kind: "date",
          code: elapsedToken[1].toLowerCase(),
          length: bracket.length,
          elapsed: true
        });
        index = end + 1;
        continue;
      }
      if (/^[hms]+$/i.test(bracket)) return void 0;
      if (!/^\$(?:[^\]]*)-[0-9a-f]+$/i.test(bracket) && !/^[a-z]+$/i.test(bracket)) return void 0;
      index = end + 1;
      continue;
    }
    if (/[dmyhs]/i.test(char)) {
      const code = char.toLowerCase();
      let end = index + 1;
      while (((_a = formatCode[end]) == null ? void 0 : _a.toLowerCase()) === code) end++;
      const length = end - index;
      if (code === "d" && length > 4 || code === "m" && length > 5 || code === "y" && length !== 2 && length !== 4 || (code === "h" || code === "s") && length > 2) return void 0;
      tokens.push({
        kind: "date",
        code,
        length
      });
      index = end;
      continue;
    }
    if (char === ".") {
      const previousToken = tokens.at(-1);
      if ((previousToken == null ? void 0 : previousToken.kind) === "date" && previousToken.code === "s") {
        const fraction = /^\.0{1,6}(?!0)/.exec(remaining);
        if (!fraction) return void 0;
        tokens.push({
          kind: "literal",
          value: "."
        });
        tokens.push({
          kind: "fraction",
          digits: fraction[0].length - 1
        });
        index += fraction[0].length;
        continue;
      }
    }
    if (/[/\-.,: ]/.test(char)) {
      tokens.push({
        kind: "literal",
        value: char
      });
      index++;
      continue;
    }
    return;
  }
  const dateTokens = tokens.filter((token) => token.kind === "date");
  if (!dateTokens.length) return void 0;
  const hasClock = dateTokens.some((token) => token.code === "h" || token.code === "s") || tokens.some((token) => token.kind === "ampm");
  return tokens.map((token, index) => {
    if (token.kind !== "date" || token.code !== "m" || token.elapsed) return token;
    const previous = tokens.slice(0, index).reverse().find((item) => item.kind !== "literal" || !/^[: ]+$/.test(item.value));
    const next = tokens.slice(index + 1).find((item) => item.kind !== "literal" || !/^[: ]+$/.test(item.value));
    return hasClock && ((previous == null ? void 0 : previous.kind) === "date" && previous.code === "h" || (next == null ? void 0 : next.kind) === "date" && next.code === "s") ? {
      ...token,
      length: -token.length
    } : token;
  });
}
function formatDateNumber(value, tokens, date1904) {
  var _a;
  if (!Number.isFinite(value) || value < 0 || value > 2958465) return void 0;
  const fractionDigits = (_a = tokens.find((token) => token.kind === "fraction")) == null ? void 0 : _a.digits;
  const precisionDigits = Math.max(fractionDigits != null ? fractionDigits : 3, 3);
  const unitsPerSecond = 10 ** precisionDigits;
  const unitsPerDay = 86400 * unitsPerSecond;
  const displayStep = 10 ** (precisionDigits - (fractionDigits != null ? fractionDigits : 3));
  let serialDay = Math.floor(value);
  let timeUnits = Math.round((value - serialDay) * unitsPerDay / displayStep) * displayStep;
  if (timeUnits >= unitsPerDay) {
    serialDay++;
    timeUnits = 0;
  }
  const fictitiousLeapDay = !date1904 && serialDay === 60;
  const adjustedDay = !date1904 && serialDay > 60 ? serialDay - 1 : serialDay;
  const date = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)) + adjustedDay * 864e5 + Math.floor(timeUnits * 1e3 / unitsPerSecond));
  if (!Number.isFinite(date.getTime())) return void 0;
  const year = date.getUTCFullYear();
  const month = fictitiousLeapDay ? 1 : date.getUTCMonth();
  const day = fictitiousLeapDay ? 29 : date.getUTCDate();
  const weekdayDate = fictitiousLeapDay ? new Date(Date.UTC(1900, 2, 1)) : date;
  const hour = date.getUTCHours();
  const minute = date.getUTCMinutes();
  const second = date.getUTCSeconds();
  const fractionValue = Math.floor(timeUnits % unitsPerSecond / displayStep);
  const totalSeconds = serialDay * 86400 + timeUnits / unitsPerSecond;
  const ampm = tokens.some((token) => token.kind === "ampm");
  const monthDate = fictitiousLeapDay ? new Date(Date.UTC(1900, 1, 1)) : date;
  const local = (options, target = date) => new Intl.DateTimeFormat(void 0, {
    ...options,
    timeZone: "UTC"
  }).format(target);
  return tokens.map((token) => {
    if (token.kind === "literal") return token.value;
    if (token.kind === "ampm") return token.short ? hour < 12 ? "A" : "P" : hour < 12 ? "AM" : "PM";
    if (token.kind === "fraction") return String(fractionValue).padStart(token.digits, "0");
    if (token.elapsed) {
      const unitSeconds = token.code === "h" ? 3600 : token.code === "m" ? 60 : 1;
      const duration = String(Math.floor(totalSeconds / unitSeconds));
      return token.length === 1 ? duration : duration.padStart(2, "0");
    }
    if (token.length < 0) return -token.length === 1 ? String(minute) : String(minute).padStart(2, "0");
    if (token.code === "d") {
      if (token.length === 1) return String(day);
      if (token.length === 2) return String(day).padStart(2, "0");
      return local({ weekday: token.length === 3 ? "short" : "long" }, weekdayDate);
    }
    if (token.code === "m") {
      if (token.length === 1) return String(month + 1);
      if (token.length === 2) return String(month + 1).padStart(2, "0");
      if (token.length === 3) return local({ month: "short" }, monthDate);
      if (token.length === 4) return local({ month: "long" }, monthDate);
      return local({ month: "narrow" }, monthDate);
    }
    if (token.code === "y") return token.length === 2 ? String(year % 100).padStart(2, "0") : String(year);
    if (token.code === "h") {
      const value2 = ampm ? hour % 12 || 12 : hour;
      return token.length === 1 ? String(value2) : String(value2).padStart(2, "0");
    }
    return token.length === 1 ? String(second) : String(second).padStart(2, "0");
  }).join("");
}
function splitNumberFormatSections(formatCode) {
  const sections = [];
  let start = 0;
  let quoted = false;
  for (let index = 0; index < formatCode.length; index++) {
    const char = formatCode[index];
    if (char === "\\" || !quoted && (char === "_" || char === "*")) {
      index++;
      continue;
    }
    if (char === '"') quoted = !quoted;
    else if (char === ";" && !quoted) {
      sections.push(formatCode.slice(start, index));
      start = index + 1;
    }
  }
  if (quoted) return void 0;
  sections.push(formatCode.slice(start));
  return sections;
}
function parseNumberFormatLiteral(value) {
  let literal = "";
  let quoted = false;
  for (let index = 0; index < value.length; index++) {
    const char = value[index];
    if (char === "\\") {
      if (index + 1 >= value.length) return void 0;
      literal += value[++index];
    } else if (char === '"') quoted = !quoted;
    else if (quoted) literal += char;
    else if (char === "_") {
      if (index + 1 >= value.length) return void 0;
      literal += " ";
      index++;
    } else if (/^[ $+():^'{}<>=\/!&~€£¥¢₹₽-]$/.test(char)) literal += char;
    else return void 0;
  }
  return quoted ? void 0 : literal;
}
function formatFractionMaskInteger(value, mask) {
  const minimumIntegerDigits = Math.max(1, [...mask].filter((char) => char === "0").length);
  const formatted = new Intl.NumberFormat(void 0, {
    useGrouping: mask.includes(","),
    minimumIntegerDigits,
    maximumFractionDigits: 0
  }).format(value);
  return mask.includes("?") ? formatted.padStart(mask.length, "\xA0") : formatted;
}
function greatestCommonDivisor(left, right) {
  while (right) [left, right] = [right, left % right];
  return left;
}
function formatFractionNumber(value, fraction) {
  if (!Number.isFinite(value)) return void 0;
  const magnitude = Math.abs(value) * (fraction.percent ? 100 : 1);
  const hasWhole = Boolean(fraction.whole);
  let whole = hasWhole ? Math.floor(magnitude) : 0;
  if (!Number.isSafeInteger(whole)) return void 0;
  const target = hasWhole ? magnitude - whole : magnitude;
  let numerator = 0;
  let denominator = fraction.fixedDenominator || 1;
  if (fraction.fixedDenominator) {
    if (target > Number.MAX_SAFE_INTEGER / denominator) return void 0;
    numerator = Math.round(target * denominator);
  } else {
    if (target > Number.MAX_SAFE_INTEGER / fraction.maximumDenominator) return void 0;
    let bestError = Number.POSITIVE_INFINITY;
    for (let candidate = 1; candidate <= fraction.maximumDenominator; candidate++) {
      const candidateNumerator = Math.round(target * candidate);
      const error = Math.abs(target - candidateNumerator / candidate);
      if (error < bestError - Number.EPSILON * Math.max(1, target)) {
        bestError = error;
        numerator = candidateNumerator;
        denominator = candidate;
      }
    }
    if (numerator) {
      const divisor = greatestCommonDivisor(numerator, denominator);
      numerator /= divisor;
      denominator /= divisor;
    }
  }
  if (hasWhole && numerator >= denominator) {
    whole++;
    numerator = 0;
  }
  const wholeMask = fraction.whole;
  const wholeText = Boolean(wholeMask && (whole > 0 || wholeMask.includes("0"))) ? formatFractionMaskInteger(whole, wholeMask) : "";
  const sign = value < 0 ? "-" : "";
  if (!numerator) return `${sign}${wholeText || "0"}`;
  const fractionText = `${formatFractionMaskInteger(numerator, fraction.numerator)}/${fraction.fixedDenominator ? new Intl.NumberFormat(void 0, { maximumFractionDigits: 0 }).format(denominator) : formatFractionMaskInteger(denominator, fraction.denominator)}`;
  return `${sign}${wholeText ? `${wholeText} ` : ""}${fractionText}`;
}
function parseNumberFormatSection(code, index) {
  var _a;
  if (!code) return null;
  const parentheses = code.startsWith("(") && code.endsWith(")");
  if (parentheses && index !== 1) return void 0;
  if (parentheses) code = code.slice(1, -1);
  let start = -1;
  let quoted = false;
  for (let offset = 0; offset < code.length; offset++) {
    const char = code[offset];
    if (char === "\\" || !quoted && char === "_") {
      offset++;
      continue;
    }
    if (char === '"') quoted = !quoted;
    else if (!quoted && (char === "#" || char === "0" || char === "?")) {
      start = offset;
      break;
    }
  }
  if (start < 0) return void 0;
  const fractionMatch = /^(?:([#0?]+(?:,[#0?]{3})*) +)?([#0?]+)\/([#0?]+|[1-9]\d*)/.exec(code.slice(start));
  if (fractionMatch) {
    const whole = fractionMatch[1] || void 0;
    const numerator = fractionMatch[2];
    const denominator = fractionMatch[3];
    const fixedDenominator = /^\d+$/.test(denominator) ? Number(denominator) : void 0;
    const denominatorEnd = start + fractionMatch[0].length;
    const rawSuffix = code.slice(denominatorEnd);
    const percentMatch = /^(\s*)%/.exec(rawSuffix);
    const percent = Boolean(percentMatch);
    const suffixStart = denominatorEnd + ((percentMatch == null ? void 0 : percentMatch[0].length) || 0);
    const prefix2 = parseNumberFormatLiteral(code.slice(0, start));
    const percentSpacing = percentMatch ? parseNumberFormatLiteral(percentMatch[1]) : "";
    const remainingSuffix = parseNumberFormatLiteral(code.slice(suffixStart));
    if (prefix2 === void 0 || percentSpacing === void 0 || remainingSuffix === void 0 || !percent && rawSuffix.includes("%") || numerator.length > 3 || fixedDenominator === void 0 && denominator.length > 3 || whole && whole.replaceAll(",", "").length > 15 || fixedDenominator !== void 0 && (!Number.isSafeInteger(fixedDenominator) || fixedDenominator < 1 || fixedDenominator > 999999)) return void 0;
    const suffix2 = `${percentSpacing}${percent ? "%" : ""}${remainingSuffix}`;
    return {
      options: {
        useGrouping: Boolean(whole == null ? void 0 : whole.includes(",")),
        minimumIntegerDigits: 1,
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
      },
      parentheses,
      prefix: prefix2,
      suffix: suffix2,
      scale: 0,
      fraction: {
        whole,
        numerator,
        denominator,
        ...fixedDenominator !== void 0 ? { fixedDenominator } : {},
        maximumDenominator: fixedDenominator === void 0 ? 10 ** denominator.length - 1 : 0,
        percent
      }
    };
  }
  let end = start;
  while (/[#0,.%]/.test(code[end] || "")) end++;
  const maskEnd = end;
  let scientific;
  if (/[Ee]/.test(code[end] || "")) {
    const exponent = /^([Ee])([+-])(0+)/.exec(code.slice(end));
    if (!exponent) return void 0;
    const integerDigits = code.slice(start, end).split(".")[0].replaceAll(",", "").replaceAll("%", "").length;
    scientific = {
      symbol: exponent[1],
      sign: exponent[2],
      minimumDigits: exponent[3].length,
      integerDigits
    };
    end += exponent[0].length;
  }
  let mask = code.slice(start, maskEnd);
  const trailingPercent = Boolean(scientific && code[end] === "%");
  if (trailingPercent) end++;
  const percentCount = [...mask].filter((char) => char === "%").length + Number(trailingPercent);
  if (percentCount > 1) return void 0;
  mask = mask.replace("%", "");
  let scale = 0;
  while (mask.endsWith(",")) {
    scale++;
    mask = mask.slice(0, -1);
  }
  if (scientific && scale) return void 0;
  const match = /^([#0]+(?:,[#0]{3})*)(?:\.([#0]+))?$/.exec(mask);
  if (!match) return void 0;
  if ((((_a = match[2]) == null ? void 0 : _a.length) || 0) > 100) return void 0;
  const prefix = parseNumberFormatLiteral(code.slice(0, start));
  const suffix = parseNumberFormatLiteral(code.slice(end));
  if (prefix === void 0 || suffix === void 0) return void 0;
  const fraction = match[2] || "";
  return {
    options: {
      useGrouping: match[1].includes(","),
      minimumIntegerDigits: Math.max(1, [...match[1]].filter((char) => char === "0").length),
      minimumFractionDigits: [...fraction].filter((char) => char === "0").length,
      maximumFractionDigits: fraction.length,
      ...percentCount ? { style: "percent" } : {}
    },
    parentheses,
    prefix,
    suffix,
    scale,
    ...scientific ? { scientific } : {}
  };
}
function parseNumberFormat(formatCode) {
  const codes = splitNumberFormatSections(formatCode);
  if (!codes || codes.length > 3 || !codes[0]) return void 0;
  const sections = codes.map(parseNumberFormatSection);
  return sections.some((section) => section === void 0) ? void 0 : sections;
}
function formatScientificNumber(value, section) {
  var _a;
  const scientific = section.scientific;
  const integerDigits = scientific.integerDigits;
  if (!Number.isFinite(value)) return String(value);
  let exponent = 0;
  let mantissa = 0;
  if (value !== 0) {
    const [coefficient, rawExponentText] = Math.abs(value).toExponential().split("e");
    const rawExponent = Number(rawExponentText) + (section.options.style === "percent" ? 2 : 0);
    exponent = Math.floor(rawExponent / integerDigits) * integerDigits;
    mantissa = Number(coefficient) * 10 ** (rawExponent - exponent);
  }
  const roundingScale = 10 ** section.options.maximumFractionDigits;
  if (Math.round(mantissa * roundingScale) / roundingScale >= 10 ** integerDigits) {
    exponent += integerDigits;
    mantissa /= 10 ** integerDigits;
  }
  const formatted = new Intl.NumberFormat(void 0, {
    ...section.options,
    style: "decimal"
  }).format(Math.sign(value) * mantissa);
  const exponentSign = exponent < 0 ? "-" : scientific.sign === "+" ? "+" : "";
  return `${formatted}${`${scientific.symbol}${exponentSign}${String(Math.abs(exponent)).padStart(scientific.minimumDigits, "0")}`}${section.options.style === "percent" ? ((_a = new Intl.NumberFormat(void 0, {
    style: "percent",
    maximumFractionDigits: 0
  }).formatToParts(0.01).find((part) => part.type === "percentSign")) == null ? void 0 : _a.value) || "%" : ""}`;
}
function formatChartNumber(value, formatCode, date1904 = false) {
  var _a;
  if (!formatCode || formatCode.toLowerCase() === "general") return String(value);
  const sections = splitNumberFormatSections(formatCode);
  const dateSection = sections && (value < 0 ? sections[1] : value === 0 ? sections[2] || sections[0] : sections[0]);
  const dateTokens = dateSection ? parseDateFormat(dateSection) : void 0;
  if (dateTokens) return (_a = formatDateNumber(Math.abs(value), dateTokens, date1904)) != null ? _a : String(value);
  const numberSections = parseNumberFormat(formatCode);
  if (!numberSections) return String(value);
  let section = numberSections[0];
  let formattedValue = value;
  if (value < 0 && numberSections.length > 1) {
    const negative = numberSections[1];
    if (!negative) return "";
    section = negative;
    formattedValue = Math.abs(value);
  } else if (value === 0 && numberSections.length > 2) {
    const zero = numberSections[2];
    if (!zero) return "";
    section = zero;
  }
  let scaledValue = formattedValue;
  for (let index = 0; index < section.scale; index++) scaledValue /= 1e3;
  const formatted = section.fraction ? formatFractionNumber(scaledValue, section.fraction) : section.scientific ? formatScientificNumber(scaledValue, section) : new Intl.NumberFormat(void 0, section.options).format(scaledValue);
  if (formatted === void 0) return String(value);
  const valueText = `${section.prefix}${formatted}${section.suffix}`;
  return section.parentheses ? `(${valueText})` : valueText;
}
function alphabeticNumber(value, alphabet = "abcdefghijklmnopqrstuvwxyz") {
  let number = Math.max(1, value);
  let result = "";
  while (number > 0) {
    number--;
    result = alphabet.charAt(number % alphabet.length) + result;
    number = Math.floor(number / alphabet.length);
  }
  return result;
}
function romanNumber(value) {
  let number = Math.max(1, value);
  let result = "";
  for (const [amount, numeral] of [
    [1e3, "m"],
    [900, "cm"],
    [500, "d"],
    [400, "cd"],
    [100, "c"],
    [90, "xc"],
    [50, "l"],
    [40, "xl"],
    [10, "x"],
    [9, "ix"],
    [5, "v"],
    [4, "iv"],
    [1, "i"]
  ]) while (number >= amount) {
    result += numeral;
    number -= amount;
  }
  return result;
}
function scriptDigits(value, digits) {
  return String(value).replace(/\d/g, (digit) => digits.charAt(Number(digit)));
}
function chineseNumber(value, traditional) {
  if (!Number.isSafeInteger(value) || value <= 0) return String(value);
  const digits = "\u96F6\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D";
  const units = traditional ? [
    "",
    "\u842C",
    "\u5104",
    "\u5146",
    "\u4EAC"
  ] : [
    "",
    "\u4E07",
    "\u4EBF",
    "\u5146",
    "\u4EAC"
  ];
  const groupUnits = [
    "",
    "\u5341",
    "\u767E",
    "\u5343"
  ];
  const groups = [];
  for (let remaining = value; remaining > 0; remaining = Math.floor(remaining / 1e4)) groups.unshift(remaining % 1e4);
  let result = "";
  let skippedGroup = false;
  groups.forEach((group, index) => {
    if (!group) {
      if (result) skippedGroup = true;
      return;
    }
    if (result && (skippedGroup || group < 1e3)) result += digits[0];
    const omitLeadingTen = !result && group >= 10 && group < 20;
    let section = "";
    let pendingZero = false;
    for (let position = 3; position >= 0; position--) {
      const digit = Math.floor(group / 10 ** position) % 10;
      if (!digit) {
        if (section) pendingZero = true;
        continue;
      }
      if (pendingZero) section += digits[0];
      if (!(omitLeadingTen && position === 1 && digit === 1 && !section)) section += digits[digit];
      section += groupUnits[position];
      pendingZero = false;
    }
    result += section + units[groups.length - index - 1];
    skippedGroup = false;
  });
  return result;
}
function numberedMarker(value, type) {
  if (type === "arabicPlain") return String(value);
  if (type === "arabicDbPlain") return scriptDigits(value, "\uFF10\uFF11\uFF12\uFF13\uFF14\uFF15\uFF16\uFF17\uFF18\uFF19");
  let label = String(value);
  if (type.startsWith("ea1Chs")) label = chineseNumber(value, false);
  else if (type.startsWith("ea1Cht")) label = chineseNumber(value, true);
  else if (type.startsWith("thaiAlpha")) label = alphabeticNumber(value, "\u0E01\u0E02\u0E04\u0E06\u0E07\u0E08\u0E09\u0E0A\u0E0B\u0E0C\u0E0D\u0E0E\u0E0F\u0E10\u0E11\u0E12\u0E13\u0E14\u0E15\u0E16\u0E17\u0E18\u0E19\u0E1A\u0E1B\u0E1C\u0E1D\u0E1E\u0E1F\u0E20\u0E21\u0E22\u0E23\u0E25\u0E27\u0E28\u0E29\u0E2A\u0E2B\u0E2C\u0E2D\u0E2E");
  else if (type.startsWith("thaiNum")) label = scriptDigits(value, "\u0E50\u0E51\u0E52\u0E53\u0E54\u0E55\u0E56\u0E57\u0E58\u0E59");
  else if (type.startsWith("hindiNum")) label = scriptDigits(value, "\u0966\u0967\u0968\u0969\u096A\u096B\u096C\u096D\u096E\u096F");
  else if (type === "arabicDbPeriod") return `${scriptDigits(value, "\uFF10\uFF11\uFF12\uFF13\uFF14\uFF15\uFF16\uFF17\uFF18\uFF19")}\uFF0E`;
  else if (type.startsWith("alpha")) {
    label = alphabeticNumber(value);
    if (type.startsWith("alphaUc")) label = label.toUpperCase();
  } else if (type.startsWith("roman")) {
    label = romanNumber(value);
    if (type.startsWith("romanUc")) label = label.toUpperCase();
  }
  if (type === "ea1ChsPlain" || type === "ea1ChtPlain") return label;
  if (type.endsWith("ParenBoth")) return `(${label})`;
  if (type.endsWith("ParenR")) return `${label})`;
  return `${label}.`;
}
function paragraphEntries(paragraphs) {
  const counters = /* @__PURE__ */ new Map();
  return paragraphs.map((paragraph) => {
    var _a, _b;
    const bullet = paragraph.bullet;
    if (!bullet || bullet.kind !== "number") {
      counters.clear();
      return {
        paragraph,
        marker: (bullet == null ? void 0 : bullet.kind) === "char" ? bullet.char : ""
      };
    }
    const key = `${bullet.type}:${paragraph.level || 0}`;
    const number = (_b = (_a = counters.get(key)) != null ? _a : bullet.startAt) != null ? _b : 1;
    counters.set(key, number + 1);
    return {
      paragraph,
      marker: numberedMarker(number, bullet.type)
    };
  });
}
function hslToRgb(hue, saturation, lightness) {
  hue = (hue % 1 + 1) % 1;
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x = chroma * (1 - Math.abs(hue * 6 % 2 - 1));
  const match = lightness - chroma / 2;
  const [r, g, b] = hue < 1 / 6 ? [
    chroma,
    x,
    0
  ] : hue < 2 / 6 ? [
    x,
    chroma,
    0
  ] : hue < 3 / 6 ? [
    0,
    chroma,
    x
  ] : hue < 4 / 6 ? [
    0,
    x,
    chroma
  ] : hue < 5 / 6 ? [
    x,
    0,
    chroma
  ] : [
    chroma,
    0,
    x
  ];
  return [
    (r + match) * 255,
    (g + match) * 255,
    (b + match) * 255
  ];
}
var eastAsian = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303f\uff00-\uffef]/u;
var complex = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Thai}\p{Script=Lao}\p{Script=Tibetan}\p{Script=Myanmar}\p{Script=Khmer}\p{Script=Mongolian}]/u;
function characterScript(character, previous = "latin") {
  return eastAsian.test(character) ? "eastAsia" : complex.test(character) ? "complexScript" : /\p{Mark}/u.test(character) ? previous : "latin";
}
function scriptSegments(text) {
  const segments = [];
  let previous = "latin";
  for (const character of text) {
    const script = characterScript(character, previous);
    const last = segments.at(-1);
    if ((last == null ? void 0 : last.script) === script) last.text += character;
    else segments.push({
      text: character,
      script
    });
    previous = script;
  }
  return segments;
}
function normalizedFontFamily(fontFamily) {
  return (fontFamily || "").replace(/[^\p{L}\p{N} ._-]/gu, "").trim().slice(0, 128);
}
function cssFontFamily(fontFamily) {
  const safeFontFamily = normalizedFontFamily(fontFamily);
  return safeFontFamily ? `'${safeFontFamily}', Arial, sans-serif` : "Arial, sans-serif";
}
function getPptxChartStacking(series, grouping) {
  const categoryCount = Math.max(0, ...series.map((item) => item.values.length));
  const segments = series.map((item) => Array(item.values.length));
  let min = 0;
  let max = 0;
  for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
    const values = series.map((item) => item.values[categoryIndex]);
    const positiveTotal = values.reduce((total, value) => total + (value !== null && value !== void 0 && value > 0 ? value : 0), 0);
    const negativeTotal = values.reduce((total, value) => total + (value !== null && value !== void 0 && value < 0 ? -value : 0), 0);
    let positive = 0;
    let negative = 0;
    values.forEach((value, seriesIndex) => {
      if (value === null || value === void 0) return;
      const segmentValue = grouping === "percentStacked" ? value / (value < 0 ? negativeTotal || 1 : positiveTotal || 1) : value;
      if (value < 0) {
        const start = negative;
        negative += segmentValue;
        segments[seriesIndex][categoryIndex] = {
          start,
          end: negative
        };
      } else {
        const start = positive;
        positive += segmentValue;
        segments[seriesIndex][categoryIndex] = {
          start,
          end: positive
        };
      }
    });
    min = Math.min(min, negative);
    max = Math.max(max, positive);
  }
  return {
    segments,
    min,
    max
  };
}
function presetEllipsePoint(width, height, angle) {
  const radians = angle * Math.PI / 108e5;
  const ellipseAngle = Math.atan2(width * Math.sin(radians), height * Math.cos(radians));
  return [500 + 500 * Math.cos(ellipseAngle), 500 + 500 * Math.sin(ellipseAngle)];
}
function timingProgress(curve, progress) {
  const acceleration = curve.acceleration || 0;
  const deceleration = curve.deceleration || 0;
  if (!acceleration && !deceleration) return progress;
  const area = 1 - (acceleration + deceleration) / 2;
  if (acceleration && progress < acceleration) return progress ** 2 / (2 * acceleration * area);
  if (deceleration && progress > 1 - deceleration) {
    const tail = progress - (1 - deceleration);
    return (1 - deceleration - acceleration / 2 + tail - tail ** 2 / (2 * deceleration)) / area;
  }
  return (acceleration / 2 + progress - acceleration) / area;
}
function gridTransitionPathData(width, height, orientation, progress, checker) {
  const unit = Math.max(1, Math.min(width, height)) / 8;
  const columns = Math.max(1, Math.round(width / unit));
  const rows = Math.max(1, Math.round(height / unit));
  const cellWidth = width / columns;
  const cellHeight = height / rows;
  const paths = [];
  for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
    let amount = progress;
    if (checker) {
      const count = orientation === "horz" ? columns : rows;
      const position = orientation === "horz" ? column : row;
      const step = (orientation === "horz" ? row : column) % 2 === 1 ? count - 1 - position : position;
      amount = Math.max(0, Math.min(1, (progress - step / Math.max(1, count - 1) * 0.72) / 0.28));
    }
    const left = column * cellWidth - 0.5;
    const right = (column + 1) * cellWidth + 0.5;
    const top = row * cellHeight - 0.5;
    const bottom = (row + 1) * cellHeight + 0.5;
    const halfWidth = checker || orientation === "vert" ? cellWidth * amount / 2 : cellWidth / 2;
    const halfHeight = checker || orientation === "horz" ? cellHeight * amount / 2 : cellHeight / 2;
    const centerX = (left + right) / 2;
    const centerY = (top + bottom) / 2;
    paths.push(`M${(centerX - halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} L${(centerX - halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} Z`);
  }
  return paths.join(" ");
}
function gridTransitionClipPath(width, height, orientation, progress, checker) {
  return `path("${gridTransitionPathData(width, height, orientation, progress, checker)}")`;
}
function gridTransitionMaskImage(width, height, orientation, progress, checker) {
  const maskWidth = Math.max(1, width);
  const maskHeight = Math.max(1, height);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${maskWidth}" height="${maskHeight}" viewBox="0 0 ${maskWidth} ${maskHeight}" preserveAspectRatio="none"><path fill="white" d="${gridTransitionPathData(maskWidth, maskHeight, orientation, progress, checker)}"/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}
function randomBarClipPath(width, height, orientation, progress, seed = "slide") {
  const amount = Math.max(0, Math.min(progress, 1));
  const axisLength = orientation === "horizontal" ? height : width;
  const count = Math.max(1, Math.min(64, Math.ceil(axisLength / Math.max(1, Math.min(width, height) / 8))));
  const bandSize = axisLength / count;
  let state = 2166136261;
  for (const character of `${seed}:${orientation}`) state = Math.imul(state ^ character.charCodeAt(0), 16777619);
  const order = Array.from({ length: count }, (_, index) => index);
  for (let index = count - 1; index > 0; index--) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    const swap = (state >>> 0) % (index + 1);
    const current = order[index];
    order[index] = order[swap];
    order[swap] = current;
  }
  const window = 0.3;
  return `path("${order.map((rank, band) => {
    const start = rank / Math.max(1, count - 1) * 0.7;
    const reveal = Math.max(0, Math.min(1, (amount - start) / window));
    const center = (band + 0.5) * bandSize;
    const half = (bandSize * reveal + (reveal > 0 ? 1 : 0)) / 2;
    if (orientation === "horizontal") return `M0 ${(center - half).toFixed(2)} H${width} V${(center + half).toFixed(2)} H0 Z`;
    return `M${(center - half).toFixed(2)} 0 H${(center + half).toFixed(2)} V${height} H${(center - half).toFixed(2)} Z`;
  }).join(" ")}")`;
}
function clipPathMaskImage(clipPath, width, height) {
  const match = /^path\((?:(evenodd),\s*)?"([\s\S]*)"\)$/.exec(clipPath);
  if (!match) return void 0;
  const maskWidth = Math.max(1, width);
  const maskHeight = Math.max(1, height);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${maskWidth}" height="${maskHeight}" viewBox="0 0 ${maskWidth} ${maskHeight}" preserveAspectRatio="none"><path fill="white" fill-rule="${match[1] || "nonzero"}" d="${match[2] || ""}"/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}
function combClipPath(width, height, orientation, progress) {
  const amount = Math.max(0, Math.min(progress, 1));
  const axisLength = orientation === "horz" ? width : height;
  const crossAxisLength = orientation === "horz" ? height : width;
  const barCount = Math.max(1, Math.min(64, Math.ceil(crossAxisLength / Math.max(1, Math.min(width, height) / 8))));
  const barSize = crossAxisLength / barCount;
  return `path("${Array.from({ length: barCount }, (_, index) => {
    const start = index % 2 === 0 ? 0 : axisLength;
    const end = index % 2 === 0 ? axisLength * amount : axisLength * (1 - amount);
    const crossStart = index * barSize - 0.5;
    const crossEnd = (index + 1) * barSize + 0.5;
    if (orientation === "horz") return `M${Math.min(start, end).toFixed(2)} ${crossStart.toFixed(2)} H${Math.max(start, end).toFixed(2)} V${crossEnd.toFixed(2)} H${Math.min(start, end).toFixed(2)} Z`;
    return `M${crossStart.toFixed(2)} ${Math.min(start, end).toFixed(2)} H${crossEnd.toFixed(2)} V${Math.max(start, end).toFixed(2)} H${crossStart.toFixed(2)} Z`;
  }).join(" ")}")`;
}
function stripsClipPath(width, height, direction, progress) {
  const amount = Math.max(0, Math.min(progress, 1));
  const count = 8;
  const bandHeight = height / count;
  const reverseRows = direction.endsWith("u");
  const reverseColumns = direction.startsWith("l");
  const staggerWindow = 0.38;
  return `path("${Array.from({ length: count }, (_, band) => {
    const start = (reverseRows ? 7 - band : band) / 7 * 0.62;
    const reveal = Math.max(0, Math.min(1, (amount - start) / staggerWindow));
    const top = band * bandHeight - 0.5;
    const bottom = (band + 1) * bandHeight + 0.5;
    const base = reverseColumns ? width : 0;
    const travel = width * reveal;
    const slant = reveal > 0 && reveal < 1 ? bandHeight * 0.75 : 0;
    const first = reverseColumns ? base - travel : base + travel;
    const leadingTop = Math.max(0, Math.min(width, first + (reverseColumns ? -slant : slant)));
    const leadingBottom = Math.max(0, Math.min(width, first - (reverseColumns ? -slant : slant)));
    return `M${(reverseColumns ? [
      [leadingTop, top],
      [width, top],
      [width, bottom],
      [leadingBottom, bottom],
      [leadingTop, top]
    ] : [
      [0, top],
      [leadingTop, top],
      [leadingBottom, bottom],
      [0, bottom],
      [0, top]
    ]).map(([x, y]) => `${x.toFixed(2)} ${y.toFixed(2)}`).join(" L")} Z`;
  }).join(" ")}")`;
}
function barnClipPath(orientation, motion, progress) {
  const amount = Math.max(0, Math.min(progress, 1));
  const inset = (motion === "in" ? 1 - amount : amount) * 50;
  return orientation === "horizontal" ? `inset(${inset}% 0)` : `inset(0 ${inset}%)`;
}
function shapeEffectClipPath(shape, width, height, progress, direction) {
  const visible = Math.max(0, Math.min(progress, 1));
  const cx = width / 2, cy = height / 2;
  const outer = `M0 0 H${width} V${height} H0 Z`;
  if (shape === "box") {
    if (direction === "in") return `inset(${(1 - visible) * 50}% ${(1 - visible) * 50}%)`;
    if (visible === 0) return "inset(50%)";
    if (visible === 1) return "inset(0)";
    const insetX = (1 - visible) * width / 2;
    const insetY = (1 - visible) * height / 2;
    return `path(evenodd, "${outer} M${insetX} ${insetY} H${width - insetX} V${height - insetY} H${insetX} Z")`;
  }
  if (shape === "plus") {
    if (direction === "in") return visible === 0 ? "inset(50%)" : visible === 1 ? "inset(0)" : `path("${plusPath(width, height, visible)}")`;
    if (visible === 1) return "inset(0)";
    if (visible === 0) return "inset(50%)";
    return `path(evenodd, "${outer} ${plusPath(width, height, 1 - visible)}")`;
  }
  if (shape === "diamond") {
    if (direction === "in" && visible === 1) return "inset(0)";
    const radiusX = width * visible, radiusY = height * visible;
    if (visible === 0) return "inset(50%)";
    const inner = `M${cx} ${cy - radiusY} L${cx + radiusX} ${cy} L${cx} ${cy + radiusY} L${cx - radiusX} ${cy} Z`;
    return direction === "in" ? `path("${inner}")` : visible === 1 ? "inset(0)" : `path(evenodd, "${outer} ${inner}")`;
  }
  const radius = Math.hypot(width, height) / 2 * (direction === "in" ? visible : 1 - visible);
  if (radius <= 0) return direction === "in" ? "inset(50%)" : "inset(0)";
  const circle = `M${cx} ${cy - radius} A${radius} ${radius} 0 1 1 ${cx} ${cy + radius} A${radius} ${radius} 0 1 1 ${cx} ${cy - radius} Z`;
  return direction === "in" ? `path("${circle}")` : radius >= Math.hypot(width, height) / 2 ? "inset(50%)" : `path(evenodd, "${outer} ${circle}")`;
}
function rippleTransitionClipPath(width, height, direction, progress) {
  const amount = Math.max(0, Math.min(1, progress));
  const [originX, originY] = {
    center: [width / 2, height / 2],
    lu: [width, height],
    ru: [0, height],
    ld: [width, 0],
    rd: [0, 0]
  }[direction];
  const maxRadius = Math.max(...[
    [0, 0],
    [width, 0],
    [0, height],
    [width, height]
  ].map(([x, y]) => Math.hypot(x - originX, y - originY)));
  const amplitude = Math.min(width, height) * 0.025 * Math.sin(Math.PI * amount);
  const points = Array.from({ length: 64 }, (_, index) => {
    const angle = Math.PI * 2 * index / 64;
    const radius = Math.max(0, maxRadius * amount + Math.sin(angle * 6 + amount * Math.PI * 4) * amplitude);
    return `${(originX + Math.cos(angle) * radius).toFixed(2)} ${(originY + Math.sin(angle) * radius).toFixed(2)}`;
  });
  return `path("M${originX.toFixed(2)} ${originY.toFixed(2)} L${points.join(" L")} Z")`;
}
function glitterTransitionClipPath(width, height, direction, pattern, progress) {
  if (width <= 0 || height <= 0) return "inset(50%)";
  const amount = Math.max(0, Math.min(1, progress));
  const moveDirections = {
    l: [-1, 0],
    r: [1, 0],
    u: [0, -1],
    d: [0, 1],
    ld: [-Math.SQRT1_2, Math.SQRT1_2],
    lu: [-Math.SQRT1_2, -Math.SQRT1_2],
    rd: [Math.SQRT1_2, Math.SQRT1_2],
    ru: [Math.SQRT1_2, -Math.SQRT1_2]
  };
  const [moveX, moveY] = moveDirections[direction] || moveDirections.l;
  const startX = -moveX, startY = -moveY;
  const projections = [
    [0, 0],
    [width, 0],
    [0, height],
    [width, height]
  ].map(([x, y]) => x * startX + y * startY);
  const minProjection = Math.min(...projections);
  const projectionSpan = Math.max(...projections) - minProjection || 1;
  const paths = [];
  const revealTile = (cx, cy, vertices) => {
    const rank = 1 - (cx * startX + cy * startY - minProjection) / projectionSpan;
    const local = Math.max(0, Math.min(1, (amount - rank * 0.72) / 0.28));
    paths.push(`M${cx.toFixed(2)} ${cy.toFixed(2)} ${vertices.map(([dx, dy]) => `L${(cx + dx * local).toFixed(2)} ${(cy + dy * local).toFixed(2)}`).join(" ")} Z`);
  };
  if (pattern === "diamond") {
    const radius = Math.sqrt(width * height / 440);
    const rows = Math.ceil(height / radius) + 3;
    const columns = Math.ceil(width / radius) + 3;
    for (let row = -rows; row <= rows; row++) for (let column = -columns; column <= columns; column++) {
      if ((row + column) % 2 !== 0) continue;
      const cx = width / 2 + column * radius;
      const cy = height / 2 + row * radius;
      if (cx + radius < 0 || cx - radius > width || cy + radius < 0 || cy - radius > height) continue;
      revealTile(cx, cy, [
        [0, -radius],
        [radius, 0],
        [0, radius],
        [-radius, 0]
      ]);
    }
  } else {
    const side = Math.sqrt(width * height / 572);
    const halfWidth = Math.sqrt(3) * side / 2;
    const rowStep = side * 1.5;
    const columnStep = halfWidth * 2;
    const vertices = [
      [0, -side],
      [halfWidth, -side / 2],
      [halfWidth, side / 2],
      [0, side],
      [-halfWidth, side / 2],
      [-halfWidth, -side / 2]
    ];
    for (let row = -Math.ceil(height / rowStep) - 2; row <= Math.ceil(height / rowStep) + 2; row++) for (let column = -Math.ceil(width / columnStep) - 2; column <= Math.ceil(width / columnStep) + 2; column++) {
      const cx = width / 2 + columnStep * (column + row / 2);
      const cy = height / 2 + rowStep * row;
      if (cx + halfWidth < 0 || cx - halfWidth > width || cy + side < 0 || cy - side > height) continue;
      revealTile(cx, cy, vertices);
    }
  }
  return `path("${paths.join(" ")}")`;
}
function shredTransitionClipPath(width, height, pattern, progress) {
  if (width <= 0 || height <= 0) return "inset(50%)";
  const amount = Math.max(0, Math.min(1, progress));
  const columns = pattern === "strip" ? 16 : 10;
  const rows = pattern === "strip" ? 1 : 7;
  const cellWidth = width / columns;
  const cellHeight = height / rows;
  const paths = [];
  for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
    const rank = pattern === "strip" ? column / Math.max(1, columns - 1) : (row + column) / (rows + columns - 2);
    const local = Math.max(0, Math.min(1, (amount - rank * 0.68) / 0.32));
    const overlap = local ? 0.25 : 0;
    const halfWidth = cellWidth * local / 2 + overlap;
    const halfHeight = cellHeight * local / 2 + overlap;
    const centerX = (column + 0.5) * cellWidth;
    const centerY = (row + 0.5) * cellHeight;
    paths.push(`M${(centerX - halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} L${(centerX - halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} Z`);
  }
  return `path("${paths.join(" ")}")`;
}
function plusPath(width, height, amount) {
  const cx = width / 2, cy = height / 2;
  const reachX = width * amount / 2;
  const reachY = height * amount / 2;
  const arm = amount * 0.28 + Math.pow(amount, 12) * 0.72;
  const armX = width * arm / 2;
  const armY = height * arm / 2;
  return `M${cx - armX} ${cy - reachY} H${cx + armX} V${cy - armY} H${cx + reachX} V${cy + armY} H${cx + armX} V${cy + reachY} H${cx - armX} V${cy + armY} H${cx - reachX} V${cy - armY} H${cx - armX} Z`;
}
function plusTransitionClipPath(width, height, progress) {
  return `path("${plusPath(width, height, Math.max(0, Math.min(1, progress)))}")`;
}
function wedgeTransitionClipPath(width, height, progress) {
  const amount = Math.max(0, Math.min(1, progress));
  const centerX = width / 2;
  const centerY = height / 2;
  const radius = Math.hypot(width, height) / 2 + 1;
  const halfAngle = Math.PI * amount;
  const points = Array.from({ length: 13 }, (_, point) => {
    const angle = -Math.PI / 2 - halfAngle + halfAngle * 2 * point / 12;
    return `L${(centerX + Math.cos(angle) * radius).toFixed(2)} ${(centerY + Math.sin(angle) * radius).toFixed(2)}`;
  });
  return `path("M${centerX.toFixed(2)} ${centerY.toFixed(2)} ${points.join(" ")} Z")`;
}
function dissolveClipPath(width, height, progress, seed = "slide") {
  const amount = Math.max(0, Math.min(progress, 1));
  if (amount === 0) return "inset(50%)";
  if (amount === 1) return "inset(0)";
  const unit = Math.max(1, Math.min(width, height)) / 8;
  const columns = Math.max(1, Math.round(width / unit));
  const rows = Math.max(1, Math.round(height / unit));
  const cellWidth = width / columns;
  const cellHeight = height / rows;
  let seedHash = 2166136261;
  for (const character of seed) seedHash = Math.imul(seedHash ^ character.charCodeAt(0), 16777619);
  const paths = [];
  for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
    let hash = Math.imul(row + 1, 374761393) ^ Math.imul(column + 1, 668265263) ^ seedHash;
    hash = Math.imul(hash ^ hash >>> 13, 1274126177);
    const visible = ((hash ^ hash >>> 16) >>> 0) / 4294967296 < amount;
    const left = column * cellWidth - 0.5;
    const right = (column + 1) * cellWidth + 0.5;
    const top = row * cellHeight - 0.5;
    const bottom = (row + 1) * cellHeight + 0.5;
    const cx = (left + right) / 2;
    const cy = (top + bottom) / 2;
    paths.push(visible ? `M${left} ${top} L${right} ${top} L${right} ${bottom} L${left} ${bottom} Z` : `M${cx} ${cy} L${cx} ${cy} L${cx} ${cy} L${cx} ${cy} Z`);
  }
  return `path("${paths.join(" ")}")`;
}
function wheelClipPath(width, height, spokes, progress, reverse = false) {
  const centerX = width / 2;
  const centerY = height / 2;
  const radius = Math.hypot(width, height) / 2 + 1;
  const pointsPerSpoke = 12;
  return `path("${Array.from({ length: spokes }, (_, spoke) => {
    const centerAngle = -Math.PI / 2 + (spoke + 0.5) * Math.PI * 2 / spokes;
    const halfAngle = Math.PI / spokes * progress;
    const points = Array.from({ length: 13 }, (_2, point) => {
      const angle = reverse ? -Math.PI / 2 + (spoke + 1) * Math.PI * 2 / spokes - Math.PI * 2 / spokes * progress * point / pointsPerSpoke : centerAngle - halfAngle + halfAngle * 2 * point / pointsPerSpoke;
      return `L${(centerX + Math.cos(angle) * radius).toFixed(2)} ${(centerY + Math.sin(angle) * radius).toFixed(2)}`;
    });
    return `M${centerX.toFixed(2)} ${centerY.toFixed(2)} ${points.join(" ")} Z`;
  }).join(" ")}")`;
}
var EMU_PER_POINT = 12700;
var SlideStage_vue_vue_type_script_setup_true_lang_default = /* @__PURE__ */ defineComponent({
  __name: "SlideStage",
  __ssrInlineRender: true,
  props: {
    slide: {},
    width: {},
    height: {},
    number: {},
    animationStep: { default: 0 },
    thumbnail: {
      type: Boolean,
      default: false
    },
    playbackPreview: {
      type: Boolean,
      default: false
    },
    triggerPlayback: {}
  },
  emits: [
    "trigger-playback",
    "activate-hyperlink",
    "stop-all-audio"
  ],
  setup(__props, { emit: __emit }) {
    const customGradientScope = `pptx-${useId()}`.replace(/[^a-zA-Z0-9_-]/g, "_");
    function defaultLegendOrigin(position, axis) {
      if (axis === "x") return position === "left" ? 88 : position === "right" ? 780 : position === "topRight" ? 760 : 88;
      return position === "left" || position === "right" ? 105 : position === "top" ? 77 : position === "topRight" ? 92 : 530;
    }
    function chartMarkerPath(marker) {
      var _a, _b;
      const { x, y } = marker;
      const radius = (_a = marker.radius) != null ? _a : 5;
      const symbol = (_b = marker.symbol) != null ? _b : "circle";
      if (symbol === "square") return `M ${x - radius} ${y - radius} h ${radius * 2} v ${radius * 2} h ${-radius * 2} Z`;
      if (symbol === "diamond") return `M ${x} ${y - radius} L ${x + radius} ${y} L ${x} ${y + radius} L ${x - radius} ${y} Z`;
      if (symbol === "triangle") return `M ${x} ${y - radius} L ${x + radius} ${y + radius} L ${x - radius} ${y + radius} Z`;
      if (symbol === "dash") return `M ${x - radius} ${y - radius * 0.22} h ${radius * 2} v ${radius * 0.44} h ${-radius * 2} Z`;
      if (symbol === "plus") {
        const arm = radius * 0.28;
        return `M ${x - arm} ${y - radius} H ${x + arm} V ${y - arm} H ${x + radius} V ${y + arm} H ${x + arm} V ${y + radius} H ${x - arm} V ${y + arm} H ${x - radius} V ${y - arm} H ${x - arm} Z`;
      }
      if (symbol === "x") {
        const arm = radius * 0.28;
        return `M ${x - radius} ${y - radius + arm} L ${x - radius + arm} ${y - radius} L ${x} ${y - arm} L ${x + radius - arm} ${y - radius} L ${x + radius} ${y - radius + arm} L ${x + arm} ${y} L ${x + radius} ${y + radius - arm} L ${x + radius - arm} ${y + radius} L ${x} ${y + arm} L ${x - radius + arm} ${y + radius} L ${x - radius} ${y + radius - arm} L ${x - arm} ${y} Z`;
      }
      if (symbol === "star") return Array.from({ length: 10 }, (_, index) => {
        const angle = -Math.PI / 2 + index * Math.PI / 5;
        const pointRadius = radius * (index % 2 ? 0.45 : 1);
        return `${index ? "L" : "M"} ${x + Math.cos(angle) * pointRadius} ${y + Math.sin(angle) * pointRadius}`;
      }).join(" ") + " Z";
      return "";
    }
    function makeChartVisual(chart, elementId) {
      var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y, _z, _A, _B, _C, _D, _E, _F, _G, _H, _I, _J, _K, _L, _M, _N, _O, _P, _Q, _R, _S, _T, _U, _V, _W, _X, _Y, _Z, __, _$, _aa, _ba, _ca, _da, _ea, _fa, _ga, _ha, _ia, _ja, _ka, _la, _ma, _na, _oa, _pa, _qa, _ra, _sa, _ta, _ua, _va, _wa, _xa, _ya, _za, _Aa;
      const combo = chart.series.some((series) => series.chartType);
      const seriesType = (series) => series.chartType || chart.type;
      const axisEntryForSeries = (series) => {
        var _a2;
        return (_a2 = chart.valueAxes) == null ? void 0 : _a2.find((item) => item.id === series.valueAxisId);
      };
      const primaryAxisId = (_b = (_a = chart.valueAxes) == null ? void 0 : _a.find((item) => item.primary)) == null ? void 0 : _b.id;
      const axisValuesForId = (axisId) => chart.series.filter((series) => !axisId || (series.valueAxisId || primaryAxisId) === axisId).flatMap((series) => series.values.filter((value) => value !== null && value !== void 0));
      const scatter = chart.type === "scatter";
      const bubble = chart.type === "bubble";
      const radar = chart.type === "radar";
      const stock = chart.type === "stock" ? chart.stock : void 0;
      const xyChart = scatter || bubble;
      const position = (_c = chart.legend) == null ? void 0 : _c.position;
      const plotPosition = ((_d = chart.legend) == null ? void 0 : _d.overlay) ? void 0 : position;
      const plot = plotPosition === "left" ? {
        left: 330,
        top: 75,
        width: 608,
        height: 390
      } : plotPosition === "right" ? {
        left: 88,
        top: 75,
        width: 650,
        height: 390
      } : plotPosition === "top" ? {
        left: 88,
        top: 145,
        width: 850,
        height: 320
      } : {
        left: 88,
        top: 75,
        width: 850,
        height: 390
      };
      const radarCenterX = plot.left + plot.width / 2;
      const radarCenterY = plot.top + plot.height / 2;
      const radarRadiusX = plot.width * 0.38;
      const radarRadiusY = plot.height * 0.38;
      const radarPoint = (index, fraction = 1, padding = 0) => {
        const angle = -Math.PI / 2 + Math.PI * 2 * index / Math.max(chart.categories.length, 1);
        return {
          x: radarCenterX + Math.cos(angle) * (radarRadiusX * fraction + padding),
          y: radarCenterY + Math.sin(angle) * (radarRadiusY * fraction + padding),
          angle
        };
      };
      const values = stock ? [
        stock.highIndex,
        stock.lowIndex,
        stock.closeIndex,
        stock.openIndex
      ].filter((index) => index !== void 0).flatMap((index) => {
        var _a2;
        return ((_a2 = chart.series[index]) == null ? void 0 : _a2.values.filter((value) => value !== null && value !== void 0)) || [];
      }) : combo ? axisValuesForId(primaryAxisId) : chart.series.flatMap((series) => series.values.filter((value) => value !== null && value !== void 0));
      const stacking = !combo && [
        "bar",
        "line",
        "area"
      ].includes(chart.type) && chart.grouping && chart.grouping !== "clustered" ? getPptxChartStacking(chart.series, chart.grouping) : void 0;
      const stackedSegments = /* @__PURE__ */ new Map();
      if (combo) {
        const groupIndexes = /* @__PURE__ */ new Map();
        chart.series.forEach((series, index) => {
          var _a2;
          const groupIndex = (_a2 = series.chartGroupIndex) != null ? _a2 : 0;
          groupIndexes.set(groupIndex, [...groupIndexes.get(groupIndex) || [], index]);
        });
        for (const [groupIndex, indexes] of groupIndexes) {
          const grouping = chart.series[indexes[0]].grouping || "clustered";
          if (grouping === "clustered") continue;
          const groupStacking = getPptxChartStacking(indexes.map((index) => chart.series[index]), grouping);
          indexes.forEach((index, seriesIndex) => stackedSegments.set(index, groupStacking.segments[seriesIndex]));
        }
      }
      const stackedSegmentFor = (seriesIndex) => combo ? stackedSegments.get(seriesIndex) : stacking == null ? void 0 : stacking.segments[seriesIndex];
      const axisScaleFraction = (value, minimum2, maximum, axis) => {
        if (!(axis == null ? void 0 : axis.logBase)) return (value - minimum2) / (maximum - minimum2);
        return (Math.log(value > 0 ? value : minimum2) - Math.log(minimum2)) / (Math.log(maximum) - Math.log(minimum2));
      };
      const axisMajorTickValues = (axis, minimum2, maximum) => {
        var _a2;
        if (!(axis == null ? void 0 : axis.logBase)) {
          const unit = axis == null ? void 0 : axis.majorUnit;
          return unit ? Array.from({ length: Math.min(101, Math.floor((maximum - minimum2) / unit + 1e-9) + 1) }, (_, index) => Number((minimum2 + index * unit).toPrecision(12))) : Array.from({ length: 5 }, (_, index) => minimum2 + (maximum - minimum2) * index / 4);
        }
        const logarithmBase = Math.log(axis.logBase);
        const minimumExponent = Math.log(minimum2) / logarithmBase;
        const maximumExponent = Math.log(maximum) / logarithmBase;
        const step = (_a2 = axis.majorUnit) != null ? _a2 : 1;
        const firstExponent = Math.ceil((minimumExponent - 1e-9) / step) * step;
        const tickCount = Math.min(101, Math.floor((maximumExponent - firstExponent) / step + 1e-9) + 1);
        const values2 = Array.from({ length: Math.max(0, tickCount) }, (_, index) => axis.logBase ** (firstExponent + index * step)).filter((value) => Number.isFinite(value) && value >= minimum2 * (1 - 1e-9) && value <= maximum * (1 + 1e-9));
        return values2.length ? values2 : [minimum2, maximum];
      };
      const axisBoundsForSeries = (series) => {
        var _a2, _b2, _c2, _d2, _e2, _f2, _g2;
        const axis = ((_a2 = axisEntryForSeries(series)) == null ? void 0 : _a2.axis) || chart.valueAxis;
        if (!combo && stacking) {
          const minimum3 = (_c2 = (_b2 = chart.valueAxis) == null ? void 0 : _b2.min) != null ? _c2 : stacking.min;
          return {
            axis,
            minimum: minimum3,
            upper: (_e2 = (_d2 = chart.valueAxis) == null ? void 0 : _d2.max) != null ? _e2 : stacking.max === minimum3 ? minimum3 + 1 : stacking.max
          };
        }
        const values2 = axisValuesForId(series.valueAxisId || primaryAxisId);
        const dataMin2 = values2.length ? (axis == null ? void 0 : axis.logBase) ? Math.min(...values2) : Math.min(0, ...values2) : (axis == null ? void 0 : axis.logBase) ? 1 : 0;
        const dataMax2 = values2.length ? Math.max(...values2, dataMin2) : (axis == null ? void 0 : axis.logBase) ? axis.logBase : 1;
        const minimum2 = (_f2 = axis == null ? void 0 : axis.min) != null ? _f2 : dataMin2;
        const fallbackUpper = (axis == null ? void 0 : axis.logBase) ? minimum2 * axis.logBase : minimum2 + 1;
        return {
          axis,
          minimum: minimum2,
          upper: (_g2 = axis == null ? void 0 : axis.max) != null ? _g2 : dataMax2 === minimum2 ? fallbackUpper : dataMax2
        };
      };
      const seriesValuePosition = (series, value) => {
        const { axis, minimum: axisMinimum, upper: axisUpper } = axisBoundsForSeries(series);
        const fraction = axisScaleFraction(value, axisMinimum, axisUpper, axis);
        const axisFraction2 = (axis == null ? void 0 : axis.reverse) ? 1 - fraction : fraction;
        return chart.direction === "horizontal" ? plot.left + plot.width * axisFraction2 : plot.top + plot.height * (1 - axisFraction2);
      };
      const dataMin = (_e = stacking == null ? void 0 : stacking.min) != null ? _e : xyChart && values.length ? Math.min(...values) : Math.min(0, ...values);
      const dataMax = (_f = stacking == null ? void 0 : stacking.max) != null ? _f : xyChart && values.length ? Math.max(...values) : Math.max(0, ...values);
      const minimum = (_h = (_g = chart.valueAxis) == null ? void 0 : _g.min) != null ? _h : dataMin;
      const upper = (_j = (_i = chart.valueAxis) == null ? void 0 : _i.max) != null ? _j : dataMax === dataMin ? dataMin + 1 : dataMax;
      const xValues = chart.series.flatMap((series) => (series.xValues || []).filter((value) => value !== null && value !== void 0));
      const xDataMin = xyChart && xValues.length ? Math.min(...xValues) : 0;
      const xDataMax = xyChart && xValues.length ? Math.max(...xValues) : 1;
      const xMinimum = (_l = (_k = chart.xAxis) == null ? void 0 : _k.min) != null ? _l : xDataMin === xDataMax ? xDataMin - (Math.abs(xDataMin) * 0.05 || 0.5) : xDataMin;
      const xUpper = (_n = (_m = chart.xAxis) == null ? void 0 : _m.max) != null ? _n : xDataMin === xDataMax ? xDataMax + (Math.abs(xDataMax) * 0.05 || 0.5) : xDataMax;
      const categoryCount = Math.max(chart.categories.length, 1);
      const formatCategoryAxisLabel = (text) => {
        if (!chart.categoryAxisNumberFormat || !text.trim()) return text;
        const value = Number(text);
        return Number.isFinite(value) ? formatChartNumber(value, chart.categoryAxisNumberFormat, chart.date1904) : text;
      };
      const horizontalValueAxis = chart.direction === "horizontal";
      const valueAxisSide = (_p = (_o = chart.valueAxis) == null ? void 0 : _o.position) != null ? _p : horizontalValueAxis ? "b" : "l";
      const xAxisSide = (_r = (_q = chart.xAxis) == null ? void 0 : _q.position) != null ? _r : "b";
      const categoryAxisSide = (_s = chart.categoryAxisPosition) != null ? _s : horizontalValueAxis ? "l" : "b";
      const valueTickLabelPosition = (_u = (_t = chart.valueAxis) == null ? void 0 : _t.tickLabelPosition) != null ? _u : "nextTo";
      const categoryTickLabelPosition = (_v = chart.categoryAxisTickLabelPosition) != null ? _v : "nextTo";
      const categoryAxisLabelStep = (_w = chart.categoryAxisTickLabelSkip) != null ? _w : 1;
      const categoryAxisTickStep = (_x = chart.categoryAxisTickMarkSkip) != null ? _x : 1;
      const showCategoryAxisLabel = (index) => index % categoryAxisLabelStep === 0;
      const showCategoryAxisTick = (index) => index % categoryAxisTickStep === 0;
      const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
      const dateValues = (_y = chart.dateAxis) == null ? void 0 : _y.values;
      const dateMinimum = (_A = (_z = chart.dateAxis) == null ? void 0 : _z.min) != null ? _A : (dateValues == null ? void 0 : dateValues.length) ? Math.min(...dateValues) : 0;
      const dateMaximum = (_C = (_B = chart.dateAxis) == null ? void 0 : _B.max) != null ? _C : (dateValues == null ? void 0 : dateValues.length) ? Math.max(...dateValues) : 1;
      const hasDateAxisScale = (dateValues == null ? void 0 : dateValues.length) === categoryCount && dateMaximum > dateMinimum;
      const dateAxisMajorTickValues = hasDateAxisScale && ((_D = chart.dateAxis) == null ? void 0 : _D.majorUnit) !== void 0 && chart.dateAxis.majorTimeUnit ? getChartDateAxisMajorTicks(dateMinimum, dateMaximum, chart.dateAxis.majorUnit, chart.dateAxis.majorTimeUnit, chart.date1904) : void 0;
      const dateAxisMinorTickValues = hasDateAxisScale && ((_E = chart.dateAxis) == null ? void 0 : _E.minorUnit) !== void 0 && chart.dateAxis.minorTimeUnit ? getChartDateAxisMajorTicks(dateMinimum, dateMaximum, chart.dateAxis.minorUnit, chart.dateAxis.minorTimeUnit, chart.date1904) : void 0;
      const dateAxisFraction = (value) => {
        var _a2;
        const fraction = clamp((value - dateMinimum) / (dateMaximum - dateMinimum), 0, 1);
        return ((_a2 = chart.dateAxis) == null ? void 0 : _a2.reverse) ? 1 - fraction : fraction;
      };
      const categoryFraction = (index) => {
        if (!hasDateAxisScale || (dateValues == null ? void 0 : dateValues[index]) === void 0) return (index + 0.5) / categoryCount;
        return dateAxisFraction(dateValues[index]);
      };
      const categoryX = (index) => plot.left + plot.width * categoryFraction(index);
      const categoryY = (index) => plot.top + plot.height * categoryFraction(index);
      const categoryBandSize = (index) => {
        const dimension = horizontalValueAxis ? plot.height : plot.width;
        if (!hasDateAxisScale) return dimension / categoryCount;
        const current = categoryFraction(index);
        const gaps = [index > 0 ? Math.abs(current - categoryFraction(index - 1)) : void 0, index + 1 < categoryCount ? Math.abs(categoryFraction(index + 1) - current) : void 0].filter((gap) => gap !== void 0 && gap > 0);
        return Math.max(1, dimension * (gaps.length ? Math.min(...gaps) : 1 / categoryCount));
      };
      const axisFraction = (value) => {
        var _a2;
        const fraction = axisScaleFraction(value, minimum, upper, chart.valueAxis);
        return ((_a2 = chart.valueAxis) == null ? void 0 : _a2.reverse) ? 1 - fraction : fraction;
      };
      const xAxisFraction = (value) => {
        var _a2;
        const fraction = axisScaleFraction(value, xMinimum, xUpper, chart.xAxis);
        return ((_a2 = chart.xAxis) == null ? void 0 : _a2.reverse) ? 1 - fraction : fraction;
      };
      const xPosition = (value) => plot.left + plot.width * xAxisFraction(value);
      const valuePosition = (value) => plot.top + plot.height * (1 - axisFraction(value));
      const withoutMajorTickValues = (values2, minimum2, maximum, majorValues) => values2.filter((value) => value > minimum2 + Number.EPSILON * Math.abs(minimum2) && value < maximum - Number.EPSILON * Math.abs(maximum) && !majorValues.some((major) => Math.abs(value - major) <= 1e-9 * Math.max(Math.abs(value), Math.abs(major), 1e-300)));
      const minorAxisTickValues = (majorValues, minimum2, maximum, minorUnit, logBase) => {
        const values2 = [];
        if (logBase) {
          const logarithmBase = Math.log(logBase);
          const minimumExponent = Math.log(minimum2) / logarithmBase;
          const maximumExponent = Math.log(maximum) / logarithmBase;
          if (minorUnit && Number.isFinite(minorUnit) && minorUnit > 0) {
            const firstExponent = (Math.floor(minimumExponent / minorUnit + 1e-9) + 1) * minorUnit;
            const count = Math.min(512, Math.floor((maximumExponent - firstExponent) / minorUnit + 1e-9) + 1);
            for (let index = 0; index < count; index++) values2.push(logBase ** (firstExponent + index * minorUnit));
          } else if (Number.isInteger(logBase) && logBase <= 16) for (let exponent = Math.floor(minimumExponent); exponent < maximumExponent && values2.length < 512; exponent++) {
            const decade = logBase ** exponent;
            for (let multiplier = 2; multiplier < logBase && values2.length < 512; multiplier++) values2.push(multiplier * decade);
          }
          else for (let index = 1; index < majorValues.length && values2.length < 512; index++) {
            const start = majorValues[index - 1];
            const end = majorValues[index];
            for (let part = 1; part < 5 && values2.length < 512; part++) values2.push(Math.exp(Math.log(start) + (Math.log(end) - Math.log(start)) * part / 5));
          }
        } else if (minorUnit && Number.isFinite(minorUnit) && minorUnit > 0) {
          const count = Math.min(512, Math.floor((maximum - minimum2) / minorUnit + 1e-9));
          for (let index = 1; index <= count; index++) values2.push(Number((minimum2 + index * minorUnit).toPrecision(12)));
        } else for (let index = 1; index < majorValues.length; index++) {
          const start = majorValues[index - 1];
          const end = majorValues[index];
          for (let part = 1; part < 5; part++) values2.push(start + (end - start) * part / 5);
        }
        return withoutMajorTickValues(values2, minimum2, maximum, majorValues);
      };
      const categoryCrossingFraction = chart.categoryAxisCrossesAt !== void 0 ? (clamp(chart.categoryAxisCrossesAt, 1, categoryCount) - 0.5) / categoryCount : ((_F = chart.valueAxis) == null ? void 0 : _F.crosses) === "min" ? 0 : ((_G = chart.valueAxis) == null ? void 0 : _G.crosses) === "max" ? 1 : void 0;
      const xyValueAxisCrossingX = xPosition(clamp((_K = (_H = chart.valueAxis) == null ? void 0 : _H.crossesAt) != null ? _K : ((_I = chart.valueAxis) == null ? void 0 : _I.crosses) === "min" ? xMinimum : ((_J = chart.valueAxis) == null ? void 0 : _J.crosses) === "max" ? xUpper : clamp(0, xMinimum, xUpper), xMinimum, xUpper));
      const xyXAxisCrossingY = valuePosition(clamp((_O = (_L = chart.xAxis) == null ? void 0 : _L.crossesAt) != null ? _O : ((_M = chart.xAxis) == null ? void 0 : _M.crosses) === "min" ? minimum : ((_N = chart.xAxis) == null ? void 0 : _N.crosses) === "max" ? upper : clamp(0, minimum, upper), minimum, upper));
      const valueAxisCrossingPosition = categoryCrossingFraction === void 0 ? void 0 : horizontalValueAxis ? plot.top + plot.height * categoryCrossingFraction : plot.left + plot.width * categoryCrossingFraction;
      const categoryAxisCrossingValue = (_Q = (_P = chart.valueAxis) == null ? void 0 : _P.crossesAt) != null ? _Q : chart.categoryAxisCrosses === "min" ? minimum : chart.categoryAxisCrosses === "max" ? upper : clamp(0, minimum, upper);
      const categoryAxisCrossingCoordinate = horizontalValueAxis ? plot.left + plot.width * axisFraction(clamp(categoryAxisCrossingValue, minimum, upper)) : valuePosition(clamp(categoryAxisCrossingValue, minimum, upper));
      const valueAxisLineCoordinate = xyChart ? xyValueAxisCrossingX : valueAxisCrossingPosition != null ? valueAxisCrossingPosition : horizontalValueAxis ? valueAxisSide === "t" ? plot.top : plot.top + plot.height : valueAxisSide === "r" ? plot.left + plot.width : plot.left;
      const valueAxisLine = xyChart ? {
        x1: valueAxisLineCoordinate,
        y1: plot.top,
        x2: valueAxisLineCoordinate,
        y2: plot.top + plot.height
      } : horizontalValueAxis ? {
        x1: plot.left,
        y1: valueAxisLineCoordinate,
        x2: plot.left + plot.width,
        y2: valueAxisLineCoordinate
      } : {
        x1: valueAxisLineCoordinate,
        y1: plot.top,
        x2: valueAxisLineCoordinate,
        y2: plot.top + plot.height
      };
      const tickValues = axisMajorTickValues(chart.valueAxis, minimum, upper);
      const ticks = tickValues.map((value) => {
        var _a2, _b2, _c2;
        const x = radar ? radarCenterX : plot.left + plot.width * axisFraction(value);
        const y = radar ? radarCenterY - radarRadiusY * axisFraction(value) : valuePosition(value);
        return {
          value,
          label: formatChartNumber(value, (_b2 = (_a2 = chart.valueAxis) == null ? void 0 : _a2.numberFormat) != null ? _b2 : chart.grouping === "percentStacked" ? "0%" : void 0, chart.date1904),
          showLabel: valueTickLabelPosition !== "none" && !((_c2 = chart.valueAxis) == null ? void 0 : _c2.deleted),
          x,
          y,
          labelX: radar ? radarCenterX - 10 : horizontalValueAxis ? x : valueTickLabelPosition === "high" ? plot.left + plot.width + 10 : valueTickLabelPosition === "low" ? plot.left - 10 : valueAxisLineCoordinate + (valueAxisSide === "r" ? 10 : -10),
          labelY: radar ? y + 5 : horizontalValueAxis ? valueTickLabelPosition === "high" ? plot.top + plot.height + 22 : valueTickLabelPosition === "low" ? plot.top - 10 : valueAxisLineCoordinate + (valueAxisSide === "t" ? -10 : 22) : y + 5,
          labelAnchor: radar ? "end" : horizontalValueAxis ? "middle" : valueTickLabelPosition === "high" ? "start" : valueTickLabelPosition === "low" ? "end" : valueAxisSide === "r" ? "start" : "end"
        };
      });
      const secondaryValueAxes = (chart.valueAxes || []).filter((item) => !item.primary).map(({ id, axis }) => {
        var _a2, _b2, _c2, _d2;
        const values2 = axisValuesForId(id);
        const minimum2 = (_a2 = axis.min) != null ? _a2 : values2.length ? Math.min(0, ...values2) : 0;
        const upper2 = (_b2 = axis.max) != null ? _b2 : Math.max(minimum2 + 1, 0, ...values2);
        const horizontal = horizontalValueAxis;
        const side = (_c2 = axis.position) != null ? _c2 : horizontal ? "t" : "r";
        const lineCoordinate = horizontal ? side === "t" ? plot.top : plot.top + plot.height : side === "l" ? plot.left : plot.left + plot.width;
        const axisLine = horizontal ? {
          x1: plot.left,
          y1: lineCoordinate,
          x2: plot.left + plot.width,
          y2: lineCoordinate
        } : {
          x1: lineCoordinate,
          y1: plot.top,
          x2: lineCoordinate,
          y2: plot.top + plot.height
        };
        const axisTickValues = axisMajorTickValues(axis, minimum2, upper2);
        const tickPosition = (value) => {
          const fraction = axisScaleFraction(value, minimum2, upper2, axis);
          const normalized = axis.reverse ? 1 - fraction : fraction;
          return horizontal ? plot.left + plot.width * normalized : plot.top + plot.height * (1 - normalized);
        };
        const labelPosition2 = (_d2 = axis.tickLabelPosition) != null ? _d2 : "nextTo";
        return {
          id,
          horizontal,
          line: axisLine,
          ticks: axisTickValues.map((value) => {
            const position2 = tickPosition(value);
            return {
              value,
              label: formatChartNumber(value, axis.numberFormat, chart.date1904),
              showLabel: !axis.deleted && labelPosition2 !== "none",
              x: horizontal ? position2 : lineCoordinate,
              y: horizontal ? lineCoordinate : position2,
              labelX: horizontal ? position2 : lineCoordinate + (side === "l" ? -10 : 10),
              labelY: horizontal ? lineCoordinate + (side === "t" ? -10 : 22) : position2 + 5,
              labelAnchor: horizontal ? "middle" : side === "l" ? "end" : "start"
            };
          }),
          minorTicks: axis.minorTickMark && axis.minorTickMark !== "none" ? minorAxisTickValues(axisTickValues, minimum2, upper2, axis.minorUnit, axis.logBase).map((value) => {
            const position2 = tickPosition(value);
            return {
              x: horizontal ? position2 : lineCoordinate,
              y: horizontal ? lineCoordinate : position2
            };
          }) : [],
          majorGridlines: !!axis.majorGridlines,
          ...axis.majorTickMark ? { majorTickMark: axis.majorTickMark } : {},
          ...axis.minorTickMark ? { minorTickMark: axis.minorTickMark } : {},
          ...axis.deleted ? { deleted: true } : {}
        };
      });
      const xTickValues = axisMajorTickValues(chart.xAxis, xMinimum, xUpper);
      const xTickLabelPosition = (_S = (_R = chart.xAxis) == null ? void 0 : _R.tickLabelPosition) != null ? _S : "nextTo";
      const xTicks = xyChart ? xTickValues.map((value) => {
        var _a2, _b2;
        const x = xPosition(value);
        return {
          value,
          label: formatChartNumber(value, (_a2 = chart.xAxis) == null ? void 0 : _a2.numberFormat, chart.date1904),
          showLabel: xTickLabelPosition !== "none" && !((_b2 = chart.xAxis) == null ? void 0 : _b2.deleted),
          x,
          y: xyXAxisCrossingY,
          labelX: x,
          labelY: xTickLabelPosition === "high" ? plot.top - 10 : xTickLabelPosition === "low" ? plot.top + plot.height + 22 : xyXAxisCrossingY + (xAxisSide === "t" ? -10 : 22),
          labelAnchor: "middle"
        };
      }) : [];
      const axisTicks = [];
      const radarGrid = [];
      const radarSpokes = [];
      const categoryGridlines = [];
      const addAxisTick = (axis, mark, x, y, horizontal, inwardDirection, axisId, minor = false) => {
        if (!mark || mark === "none") return;
        const length = minor ? 3 : 5;
        const start = mark === "cross" ? -length : 0;
        const end = mark === "cross" ? length : mark === "in" ? length : -length;
        axisTicks.push({
          axis,
          ...axisId ? { axisId } : {},
          ...minor ? { minor: true } : {},
          x1: horizontal ? x + start * inwardDirection : x,
          y1: horizontal ? y : y + start * inwardDirection,
          x2: horizontal ? x + end * inwardDirection : x,
          y2: horizontal ? y : y + end * inwardDirection
        });
      };
      const valueTickInwardDirection = horizontalValueAxis ? valueAxisSide === "t" ? 1 : -1 : valueAxisSide === "l" ? 1 : -1;
      if (!((_T = chart.valueAxis) == null ? void 0 : _T.deleted)) {
        if (radar) for (const tick of ticks) addAxisTick("value", (_U = chart.valueAxis) == null ? void 0 : _U.majorTickMark, radarCenterX, tick.y, true, 1);
        else for (const tick of ticks) addAxisTick("value", (_V = chart.valueAxis) == null ? void 0 : _V.majorTickMark, horizontalValueAxis ? tick.x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : tick.y, !horizontalValueAxis, valueTickInwardDirection);
        if (((_W = chart.valueAxis) == null ? void 0 : _W.minorTickMark) && chart.valueAxis.minorTickMark !== "none") for (const value of minorAxisTickValues(tickValues, minimum, upper, chart.valueAxis.minorUnit, chart.valueAxis.logBase)) {
          const x = radar ? radarCenterX : plot.left + plot.width * axisFraction(value);
          const y = radar ? radarCenterY - radarRadiusY * axisFraction(value) : valuePosition(value);
          if (radar) addAxisTick("value", chart.valueAxis.minorTickMark, radarCenterX, y, true, 1, void 0, true);
          else addAxisTick("value", chart.valueAxis.minorTickMark, horizontalValueAxis ? x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : y, !horizontalValueAxis, valueTickInwardDirection, void 0, true);
        }
      }
      if (xyChart && !((_X = chart.xAxis) == null ? void 0 : _X.deleted)) {
        const inwardDirection = xAxisSide === "t" ? 1 : -1;
        for (const tick of xTicks) addAxisTick("xValue", (_Y = chart.xAxis) == null ? void 0 : _Y.majorTickMark, tick.x, tick.y, true, inwardDirection);
        if (((_Z = chart.xAxis) == null ? void 0 : _Z.minorTickMark) && chart.xAxis.minorTickMark !== "none") for (const value of minorAxisTickValues(xTickValues, xMinimum, xUpper, chart.xAxis.minorUnit, chart.xAxis.logBase)) addAxisTick("xValue", chart.xAxis.minorTickMark, xPosition(value), xyXAxisCrossingY, true, inwardDirection, void 0, true);
      }
      for (const axis of secondaryValueAxes) {
        if (axis.deleted) continue;
        const inwardDirection = axis.horizontal ? axis.line.y1 === plot.top ? 1 : -1 : axis.line.x1 === plot.left ? 1 : -1;
        for (const tick of axis.ticks) addAxisTick("valueSecondary", axis.majorTickMark, tick.x, tick.y, !axis.horizontal, inwardDirection, axis.id);
        for (const tick of axis.minorTicks) addAxisTick("valueSecondary", axis.minorTickMark, tick.x, tick.y, !axis.horizontal, inwardDirection, axis.id, true);
      }
      const categoryTickInwardDirection = horizontalValueAxis ? categoryAxisSide === "l" ? 1 : -1 : categoryAxisSide === "t" ? 1 : -1;
      if (!xyChart && !chart.categoryAxisDeleted && chart.categoryAxisMajorTickMark && chart.categoryAxisMajorTickMark !== "none") {
        if (radar) chart.categories.forEach((_category, index) => {
          if (!showCategoryAxisTick(index)) return;
          const point = radarPoint(index);
          const start = chart.categoryAxisMajorTickMark === "cross" ? -5 : chart.categoryAxisMajorTickMark === "in" ? -5 : 0;
          const end = chart.categoryAxisMajorTickMark === "cross" ? 5 : chart.categoryAxisMajorTickMark === "in" ? 0 : 5;
          axisTicks.push({
            axis: "category",
            x1: point.x + Math.cos(point.angle) * start,
            y1: point.y + Math.sin(point.angle) * start,
            x2: point.x + Math.cos(point.angle) * end,
            y2: point.y + Math.sin(point.angle) * end
          });
        });
        else for (let index = 0; index < ((__ = dateAxisMajorTickValues == null ? void 0 : dateAxisMajorTickValues.length) != null ? __ : categoryCount); index++) {
          if (!showCategoryAxisTick(index)) continue;
          const fraction = dateAxisMajorTickValues ? dateAxisFraction(dateAxisMajorTickValues[index]) : void 0;
          const centerX = fraction === void 0 ? categoryX(index) : plot.left + plot.width * fraction;
          const centerY = fraction === void 0 ? categoryY(index) : plot.top + plot.height * fraction;
          addAxisTick("category", chart.categoryAxisMajorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : centerX, horizontalValueAxis ? centerY : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection);
        }
      }
      if (!xyChart && !chart.categoryAxisDeleted && chart.categoryAxisMinorTickMark && chart.categoryAxisMinorTickMark !== "none") {
        if (radar) chart.categories.forEach((_category, index) => {
          const angle = radarPoint(index).angle + Math.PI / chart.categories.length;
          const x = radarCenterX + Math.cos(angle) * radarRadiusX;
          const y = radarCenterY + Math.sin(angle) * radarRadiusY;
          const start = chart.categoryAxisMinorTickMark === "cross" ? -3 : chart.categoryAxisMinorTickMark === "in" ? -3 : 0;
          const end = chart.categoryAxisMinorTickMark === "cross" ? 3 : chart.categoryAxisMinorTickMark === "in" ? 0 : 3;
          axisTicks.push({
            axis: "category",
            minor: true,
            x1: x + Math.cos(angle) * start,
            y1: y + Math.sin(angle) * start,
            x2: x + Math.cos(angle) * end,
            y2: y + Math.sin(angle) * end
          });
        });
        else if (hasDateAxisScale) {
          const values2 = dateAxisMinorTickValues ? withoutMajorTickValues(dateAxisMinorTickValues, dateMinimum, dateMaximum, dateAxisMajorTickValues || []) : minorAxisTickValues(dateAxisMajorTickValues || dateValues || [], dateMinimum, dateMaximum);
          for (const value of values2) {
            const fraction = dateAxisFraction(value);
            addAxisTick("category", chart.categoryAxisMinorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : plot.left + plot.width * fraction, horizontalValueAxis ? plot.top + plot.height * fraction : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection, void 0, true);
          }
        } else for (let index = 0; index < categoryCount - 1; index++) {
          const centerX = (categoryX(index) + categoryX(index + 1)) / 2;
          const centerY = (categoryY(index) + categoryY(index + 1)) / 2;
          addAxisTick("category", chart.categoryAxisMinorTickMark, horizontalValueAxis ? categoryAxisCrossingCoordinate : centerX, horizontalValueAxis ? centerY : categoryAxisCrossingCoordinate, horizontalValueAxis, categoryTickInwardDirection, void 0, true);
        }
      }
      if (!xyChart && !radar && !chart.categoryAxisDeleted && chart.categoryAxisMajorGridlines) {
        const positions = dateAxisMajorTickValues ? dateAxisMajorTickValues.map((value) => dateAxisFraction(value)) : Array.from({ length: categoryCount }, (_, index) => categoryFraction(index));
        for (const fraction of positions) if (horizontalValueAxis) {
          const y = plot.top + plot.height * fraction;
          categoryGridlines.push({
            x1: plot.left,
            y1: y,
            x2: plot.left + plot.width,
            y2: y
          });
        } else {
          const x = plot.left + plot.width * fraction;
          categoryGridlines.push({
            x1: x,
            y1: plot.top,
            x2: x,
            y2: plot.top + plot.height
          });
        }
      }
      const categories = [];
      const bars = [];
      const areas = [];
      const lines = [];
      const markers = [];
      const slices = [];
      const legend = [];
      const dataLabels = [];
      if (combo) chart.series.forEach((series, seriesIndex) => {
        var _a2, _b2, _c2, _d2, _e2, _f2, _g2;
        if (seriesType(series) !== "area") return;
        const axis = ((_a2 = axisEntryForSeries(series)) == null ? void 0 : _a2.axis) || chart.valueAxis;
        const { minimum: minimum2, upper: upper2 } = axisBoundsForSeries(series);
        const lowerValue = (_b2 = axis == null ? void 0 : axis.crossesAt) != null ? _b2 : (axis == null ? void 0 : axis.crosses) === "max" ? upper2 : (axis == null ? void 0 : axis.crosses) === "min" ? minimum2 : clamp(0, minimum2, upper2);
        let upperPoints = [];
        let lowerPoints = [];
        const flushArea = () => {
          if (upperPoints.length > 1) {
            const points = [...upperPoints, ...lowerPoints.reverse()];
            areas.push({
              d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z`,
              color: series.color
            });
          }
          upperPoints = [];
          lowerPoints = [];
        };
        for (let index = 0; index < categoryCount; index++) {
          const value = series.values[index];
          if (value == null) {
            flushArea();
            continue;
          }
          const stackedSegment = (_c2 = stackedSegmentFor(seriesIndex)) == null ? void 0 : _c2[index];
          const centerX = categoryX(index);
          const centerY = categoryY(index);
          upperPoints.push(horizontalValueAxis ? {
            x: seriesValuePosition(series, (_d2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _d2 : value),
            y: centerY
          } : {
            x: centerX,
            y: seriesValuePosition(series, (_e2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _e2 : value)
          });
          lowerPoints.push(horizontalValueAxis ? {
            x: seriesValuePosition(series, (_f2 = stackedSegment == null ? void 0 : stackedSegment.start) != null ? _f2 : lowerValue),
            y: centerY
          } : {
            x: centerX,
            y: seriesValuePosition(series, (_g2 = stackedSegment == null ? void 0 : stackedSegment.start) != null ? _g2 : lowerValue)
          });
        }
        flushArea();
      });
      const formatValue = (value, series, index) => {
        var _a2, _b2;
        return formatChartNumber(value, (_b2 = (_a2 = series.valueNumberFormatOverrides[index]) != null ? _a2 : series.numberFormat) != null ? _b2 : chart.numberFormat, chart.date1904);
      };
      const labelPosition = (series, index, fallback) => {
        var _a2, _b2, _c2;
        return (_c2 = (_b2 = (_a2 = series.dataLabelPositionOverrides[index]) != null ? _a2 : series.dataLabelPosition) != null ? _b2 : chart.dataLabelPosition) != null ? _c2 : fallback;
      };
      const chartLegendKey = (series, index) => {
        var _a2, _b2;
        const color = index === void 0 ? series.color : series.pointColors[index] || series.color;
        const type = seriesType(series);
        if (![
          "line",
          "scatter",
          "radar"
        ].includes(type)) return { color };
        const symbol = index === void 0 ? series.markerSymbol : (_b2 = (_a2 = series.markerSymbolOverrides) == null ? void 0 : _a2[index]) != null ? _b2 : series.markerSymbol;
        const showMarker = symbol === void 0 ? type === "line" ? series.lineShowMarker !== false : type === "scatter" ? series.scatterShowMarker === true : series.radarShowMarker === true : symbol !== "none";
        const showLine = type === "scatter" ? series.scatterShowLine === true : type === "line" ? series.lineShowLine !== false : series.radarShowLine !== false;
        return {
          color,
          markerChart: true,
          ...showMarker ? { symbol: symbol != null ? symbol : "circle" } : {},
          ...showLine ? { line: true } : {}
        };
      };
      const pushDataLabel = (text, x, y, anchor, series, index) => {
        var _a2;
        const style = {
          ...series.dataLabelStyle,
          ...series.dataLabelStyleOverrides[index]
        };
        const lines2 = text.split(/\r?\n/);
        const showLegendKey = (_a2 = series.legendKeyOverrides[index]) != null ? _a2 : series.showLegendKey;
        dataLabels.push({
          text,
          lines: lines2,
          x,
          y,
          anchor,
          ...showLegendKey ? { legendKey: chartLegendKey(series, index) } : {},
          ...Object.keys(style).length ? { style } : {}
        });
      };
      const formatDataLabel = (series, index, value, category, total, bubbleSize) => {
        var _a2, _b2, _c2, _d2, _e2, _f2, _g2, _h2, _i2, _j2, _k2, _l2;
        const showValue = (_a2 = series.valueLabelOverrides[index]) != null ? _a2 : series.showValueLabels;
        const showLegendKey = (_b2 = series.legendKeyOverrides[index]) != null ? _b2 : series.showLegendKey;
        const showCategory = (_c2 = series.categoryNameLabelOverrides[index]) != null ? _c2 : series.showCategoryNameLabels;
        const showSeries = (_d2 = series.seriesNameLabelOverrides[index]) != null ? _d2 : series.showSeriesNameLabels;
        const isPieLike = chart.type === "pie" || chart.type === "doughnut";
        const showPercent = isPieLike && ((_e2 = series.percentLabelOverrides[index]) != null ? _e2 : series.showPercentLabels);
        const showBubbleSize = bubble && bubbleSize !== void 0 && ((_f2 = series.bubbleSizeLabelOverrides[index]) != null ? _f2 : series.showBubbleSizeLabels);
        const parts = [
          showValue ? formatValue(value, series, index) : void 0,
          showCategory ? category : void 0,
          showSeries ? series.name : void 0,
          showPercent && total ? formatChartNumber(value / total, (_i2 = (_h2 = (_g2 = series.valueNumberFormatOverrides[index]) != null ? _g2 : series.numberFormat) != null ? _h2 : chart.numberFormat) != null ? _i2 : "0.0%", chart.date1904) : void 0,
          showBubbleSize ? formatValue(bubbleSize, series, index) : void 0
        ].filter((part) => part !== void 0);
        if (!parts.length) return showLegendKey ? "" : void 0;
        const defaultSeparator = isPieLike && !showValue && showCategory && !showSeries && showPercent ? "\n" : ",";
        return parts.join((_l2 = (_k2 = (_j2 = series.labelSeparatorOverrides[index]) != null ? _j2 : series.labelSeparator) != null ? _k2 : chart.labelSeparator) != null ? _l2 : defaultSeparator);
      };
      if (stock) {
        const highSeries = chart.series[stock.highIndex];
        const lowSeries = chart.series[stock.lowIndex];
        const closeSeries = chart.series[stock.closeIndex];
        const openSeries = stock.openIndex === void 0 ? void 0 : chart.series[stock.openIndex];
        const highLowSegments = [];
        const dropSegments = [];
        const priceIndices = [
          stock.openIndex,
          stock.highIndex,
          stock.lowIndex,
          stock.closeIndex
        ].filter((index) => index !== void 0);
        const bodyWidthFactor = 100 / (100 + stock.gapWidth);
        const volumeSeries = stock.volumeGapWidth === void 0 || stock.volumeIndex === void 0 ? void 0 : chart.series[stock.volumeIndex];
        if (volumeSeries) {
          const { minimum: volumeMinimum, upper: volumeUpper } = axisBoundsForSeries(volumeSeries);
          const zeroY = seriesValuePosition(volumeSeries, clamp(0, volumeMinimum, volumeUpper));
          const volumeWidthFactor = 100 / (100 + ((_$ = stock.volumeGapWidth) != null ? _$ : 150));
          for (let index = 0; index < chart.categories.length; index++) {
            const value = volumeSeries.values[index];
            if (value == null) continue;
            const x = categoryX(index);
            const volumeWidth = categoryBandSize(index) * volumeWidthFactor;
            const valueY = seriesValuePosition(volumeSeries, value);
            const top = Math.min(zeroY, valueY);
            const bottom = Math.max(zeroY, valueY);
            bars.push({
              x: x - volumeWidth / 2,
              y: top,
              width: volumeWidth,
              height: Math.max(1, bottom - top),
              color: volumeSeries.pointColors[index] || volumeSeries.color
            });
            const label = formatDataLabel(volumeSeries, index, value, chart.categories[index] || String(index + 1));
            if (label !== void 0) {
              const position2 = labelPosition(volumeSeries, index, "outEnd");
              let labelX = x;
              let labelY = value >= 0 ? top - 8 : bottom + 20;
              let anchor = "middle";
              if (position2 === "ctr") labelY = (zeroY + valueY) / 2 + 6;
              else if (position2 === "inBase") labelY = value >= 0 ? zeroY - 8 : zeroY + 18;
              else if (position2 === "inEnd") labelY = value >= 0 ? valueY + 18 : valueY - 4;
              else if (position2 === "b") labelY = bottom + 20;
              else if (position2 === "l") {
                labelX -= volumeWidth / 2 + 8;
                labelY = (zeroY + valueY) / 2 + 6;
                anchor = "end";
              } else if (position2 === "r") {
                labelX += volumeWidth / 2 + 8;
                labelY = (zeroY + valueY) / 2 + 6;
                anchor = "start";
              }
              pushDataLabel(label, labelX, labelY, anchor, volumeSeries, index);
            }
          }
        }
        for (let index = 0; index < chart.categories.length; index++) {
          const highValue = highSeries == null ? void 0 : highSeries.values[index];
          const lowValue = lowSeries == null ? void 0 : lowSeries.values[index];
          const closeValue = closeSeries == null ? void 0 : closeSeries.values[index];
          const openValue = openSeries == null ? void 0 : openSeries.values[index];
          if (highValue == null || lowValue == null || closeValue == null || openSeries && openValue == null) continue;
          const x = categoryX(index);
          const bodyWidth = categoryBandSize(index) * bodyWidthFactor;
          const highY = valuePosition(Math.max(highValue, lowValue));
          const lowY = valuePosition(Math.min(highValue, lowValue));
          const closeY = valuePosition(closeValue);
          if (stock.showHighLowLines) highLowSegments.push(`M ${x} ${highY} L ${x} ${lowY}`);
          if (openSeries && openValue != null) {
            const openY = valuePosition(openValue);
            if (stock.showUpDownBars) {
              const color = closeValue >= openValue ? stock.upColor : stock.downColor;
              bars.push({
                x: x - bodyWidth / 2,
                y: Math.min(openY, closeY),
                width: bodyWidth,
                height: Math.max(1, Math.abs(closeY - openY)),
                color
              });
            }
            highLowSegments.push(`M ${x - 7} ${openY} L ${x} ${openY}`, `M ${x} ${closeY} L ${x + 7} ${closeY}`);
          } else highLowSegments.push(`M ${x - 7} ${closeY} L ${x + 7} ${closeY}`);
          if (stock.showDropLines) dropSegments.push(`M ${x} ${closeY} L ${x} ${categoryAxisCrossingCoordinate}`);
          const label = formatDataLabel(closeSeries, index, closeValue, chart.categories[index] || String(index + 1));
          if (label !== void 0) {
            const position2 = labelPosition(closeSeries, index, "r");
            pushDataLabel(label, position2 === "l" ? x - 10 : position2 === "r" ? x + 10 : x, position2 === "t" ? closeY - 12 : position2 === "b" ? closeY + 20 : closeY + 5, position2 === "l" ? "end" : position2 === "r" ? "start" : "middle", closeSeries, index);
          }
        }
        if (highLowSegments.length) lines.push({
          d: highLowSegments.join(" "),
          color: stock.highLowColor,
          width: 2
        });
        if (dropSegments.length) lines.push({
          d: dropSegments.join(" "),
          color: "#94a3b8",
          width: 1
        });
        if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== "none") {
          const categoryLabelY = categoryTickLabelPosition === "high" ? ((_aa = chart.valueAxis) == null ? void 0 : _aa.reverse) ? plot.top + plot.height + 25 : plot.top - 10 : categoryTickLabelPosition === "low" ? ((_ba = chart.valueAxis) == null ? void 0 : _ba.reverse) ? plot.top - 10 : plot.top + plot.height + 25 : categoryAxisSide === "t" ? categoryAxisCrossingCoordinate - 10 : categoryAxisCrossingCoordinate + 25;
          chart.categories.forEach((text, index) => {
            if (showCategoryAxisLabel(index)) categories.push({
              text: formatCategoryAxisLabel(text),
              x: categoryX(index),
              y: categoryLabelY,
              anchor: "middle"
            });
          });
        }
        for (const index of [...volumeSeries && stock.volumeIndex !== void 0 ? [stock.volumeIndex] : [], ...priceIndices]) {
          const series = chart.series[index];
          if (series && !((_da = (_ca = chart.legend) == null ? void 0 : _ca.hiddenEntries) == null ? void 0 : _da.includes(series.legendIndex))) legend.push({
            text: series.name,
            x: 88 + legend.length * 220,
            y: 530,
            ...chartLegendKey(series)
          });
        }
      } else if (chart.type === "area" && !combo) chart.series.forEach((series, seriesIndex) => {
        var _a2, _b2, _c2, _d2, _e2;
        let upperPoints = [];
        let lowerPoints = [];
        const flushArea = () => {
          if (upperPoints.length > 1) {
            const points = [...upperPoints, ...lowerPoints.reverse()];
            areas.push({
              d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z`,
              color: series.color
            });
          }
          upperPoints = [];
          lowerPoints = [];
        };
        for (let index = 0; index < categoryCount; index++) {
          const value = series.values[index];
          if (value === null || value === void 0) {
            flushArea();
            continue;
          }
          const x = categoryX(index);
          const segment = (_a2 = stacking == null ? void 0 : stacking.segments[seriesIndex]) == null ? void 0 : _a2[index];
          const upperValue = stacking ? (_b2 = segment == null ? void 0 : segment.end) != null ? _b2 : value : value;
          const lowerValue = stacking ? (_c2 = segment == null ? void 0 : segment.start) != null ? _c2 : 0 : categoryAxisCrossingValue;
          const upperY = valuePosition(upperValue);
          const lowerY = valuePosition(lowerValue);
          upperPoints.push({
            x,
            y: upperY
          });
          lowerPoints.push({
            x,
            y: lowerY
          });
          const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
          if (label !== void 0) {
            const position2 = labelPosition(series, index, stacking ? "ctr" : "outEnd");
            const labelY = position2 === "ctr" ? (upperY + lowerY) / 2 + 5 : position2 === "inBase" ? lowerY - 8 : position2 === "inEnd" ? upperY + 18 : position2 === "b" ? Math.max(upperY, lowerY) + 20 : Math.min(upperY, lowerY) - 8;
            pushDataLabel(label, x, labelY, "middle", series, index);
          }
        }
        flushArea();
        if (!((_e2 = (_d2 = chart.legend) == null ? void 0 : _d2.hiddenEntries) == null ? void 0 : _e2.includes(series.legendIndex))) legend.push({
          text: series.name,
          x: 88 + legend.length % 4 * 220,
          y: 530 + Math.floor(legend.length / 4) * 22,
          ...chartLegendKey(series)
        });
      });
      else if (bubble) {
        const maximumSize = chart.series.flatMap((series) => (series.bubbleSizes || []).filter((size) => size !== null && size !== void 0 && size !== 0 && (size > 0 || chart.showNegativeBubbles === true))).reduce((maximum, size) => Math.max(maximum, Math.abs(size)), 0);
        const maximumRadius = Math.min(plot.width, plot.height) * 0.08 * ((_ea = chart.bubbleScale) != null ? _ea : 100) / 100;
        chart.series.forEach((series, seriesIndex) => {
          var _a2, _b2, _c2, _d2, _e2, _f2;
          const pointCount = Math.max(((_a2 = series.xValues) == null ? void 0 : _a2.length) || 0, series.values.length, ((_b2 = series.bubbleSizes) == null ? void 0 : _b2.length) || 0);
          for (let index = 0; index < pointCount; index++) {
            const xValue = (_c2 = series.xValues) == null ? void 0 : _c2[index];
            const yValue = series.values[index];
            const bubbleSize = (_d2 = series.bubbleSizes) == null ? void 0 : _d2[index];
            if (xValue == null || yValue == null || bubbleSize == null || bubbleSize === 0 || bubbleSize < 0 && chart.showNegativeBubbles !== true || maximumSize <= 0) continue;
            const fraction = Math.abs(bubbleSize) / maximumSize;
            const radius = maximumRadius * (chart.bubbleSizeRepresents === "w" ? fraction : Math.sqrt(fraction));
            if (radius <= 0) continue;
            const x = xPosition(xValue);
            const y = valuePosition(yValue);
            const color = series.pointColors[index] || (chart.bubbleVaryColors ? chart.palette[index % chart.palette.length] || series.color : series.color);
            markers.push({
              x,
              y,
              color,
              radius
            });
            const label = formatDataLabel(series, index, yValue, String(xValue), void 0, bubbleSize);
            if (label !== void 0) {
              const labelType = labelPosition(series, index, "r");
              const labelX = labelType === "l" ? x - radius - 4 : labelType === "r" ? x + radius + 4 : x;
              const labelY = labelType === "t" ? y - radius - 4 : labelType === "b" ? y + radius + 18 : y + 5;
              pushDataLabel(label, labelX, labelY, labelType === "l" ? "end" : labelType === "r" ? "start" : "middle", series, index);
            }
          }
          if (!((_f2 = (_e2 = chart.legend) == null ? void 0 : _e2.hiddenEntries) == null ? void 0 : _f2.includes(series.legendIndex))) legend.push({
            text: series.name,
            x: 88 + seriesIndex * 220,
            y: 530,
            ...chartLegendKey(series)
          });
        });
      } else if (scatter) chart.series.forEach((series, seriesIndex) => {
        var _a2, _b2, _c2, _d2, _e2, _f2, _g2, _h2;
        let path = "";
        let segmentOpen = false;
        const pointCount = Math.max(((_a2 = series.xValues) == null ? void 0 : _a2.length) || 0, series.values.length);
        for (let index = 0; index < pointCount; index++) {
          const xValue = (_b2 = series.xValues) == null ? void 0 : _b2[index];
          const yValue = series.values[index];
          if (xValue == null || yValue == null) {
            segmentOpen = false;
            continue;
          }
          const x = xPosition(xValue);
          const y = valuePosition(yValue);
          if (series.scatterShowLine) {
            path += `${segmentOpen ? " L" : " M"} ${x} ${y}`;
            segmentOpen = true;
          }
          const pointMarkerSymbol = (_d2 = (_c2 = series.markerSymbolOverrides) == null ? void 0 : _c2[index]) != null ? _d2 : series.markerSymbol;
          if (pointMarkerSymbol === void 0 ? series.scatterShowMarker : pointMarkerSymbol !== "none") markers.push({
            x,
            y,
            color: series.color,
            symbol: pointMarkerSymbol != null ? pointMarkerSymbol : "circle",
            radius: (_f2 = (_e2 = series.markerSizeOverrides) == null ? void 0 : _e2[index]) != null ? _f2 : series.markerSize
          });
          const label = formatDataLabel(series, index, yValue, String(xValue));
          if (label !== void 0) {
            const labelType = labelPosition(series, index, "r");
            const labelX = labelType === "l" ? x - 10 : labelType === "r" ? x + 10 : x;
            const labelY = labelType === "t" ? y - 12 : labelType === "b" ? y + 24 : y + 5;
            pushDataLabel(label, labelX, labelY, labelType === "l" ? "end" : labelType === "r" ? "start" : "middle", series, index);
          }
        }
        if (path) lines.push({
          d: path,
          color: series.color
        });
        if (!((_h2 = (_g2 = chart.legend) == null ? void 0 : _g2.hiddenEntries) == null ? void 0 : _h2.includes(series.legendIndex))) legend.push({
          text: series.name,
          x: 88 + seriesIndex * 220,
          y: 530,
          ...chartLegendKey(series)
        });
      });
      else if (chart.type === "pie" || chart.type === "doughnut") {
        const centerX = plot.left + plot.width / 2;
        const centerY = plot.top + plot.height / 2;
        const radius = Math.min(plot.width, plot.height) * 0.47;
        const doughnut = chart.type === "doughnut";
        const baseInnerRadius = doughnut ? radius * ((_fa = chart.holeSize) != null ? _fa : 50) / 100 : 0;
        const ringSeries = doughnut ? chart.series : chart.series.slice(0, 1);
        const ringWidth = doughnut ? (radius - baseInnerRadius) / Math.max(ringSeries.length, 1) : radius;
        ringSeries.forEach((series, seriesIndex) => {
          var _a2;
          const innerRadius = doughnut ? baseInnerRadius + ringWidth * seriesIndex : 0;
          const outerRadius = doughnut ? innerRadius + ringWidth : radius;
          const points = chart.categories.flatMap((text, index) => {
            const value = series.values[index];
            return value != null && value > 0 ? [{
              text,
              index,
              value,
              color: series.pointColors[index] || chart.palette[index % chart.palette.length] || series.color
            }] : [];
          });
          const total = points.reduce((sum, point) => sum + point.value, 0);
          let angle = -Math.PI / 2 + ((_a2 = chart.firstSliceAngle) != null ? _a2 : 0) * Math.PI / 180;
          for (const point of points) {
            const sweep = point.value / total * Math.PI * 2;
            const startX = centerX + Math.cos(angle) * outerRadius;
            const startY = centerY + Math.sin(angle) * outerRadius;
            const endAngle = angle + sweep;
            const endX = centerX + Math.cos(endAngle) * outerRadius;
            const endY = centerY + Math.sin(endAngle) * outerRadius;
            const endInnerX = centerX + Math.cos(endAngle) * innerRadius;
            const endInnerY = centerY + Math.sin(endAngle) * innerRadius;
            const startInnerX = centerX + Math.cos(angle) * innerRadius;
            const startInnerY = centerY + Math.sin(angle) * innerRadius;
            const fullCircle = sweep >= Math.PI * 2 - 1e-4;
            const d = innerRadius > 0 ? fullCircle ? `M ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 1 1 ${centerX - Math.cos(angle) * outerRadius} ${centerY - Math.sin(angle) * outerRadius} A ${outerRadius} ${outerRadius} 0 1 1 ${startX} ${startY} L ${startInnerX} ${startInnerY} A ${innerRadius} ${innerRadius} 0 1 0 ${centerX - Math.cos(angle) * innerRadius} ${centerY - Math.sin(angle) * innerRadius} A ${innerRadius} ${innerRadius} 0 1 0 ${startInnerX} ${startInnerY} Z` : `M ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${endX} ${endY} L ${endInnerX} ${endInnerY} A ${innerRadius} ${innerRadius} 0 ${sweep > Math.PI ? 1 : 0} 0 ${startInnerX} ${startInnerY} Z` : fullCircle ? `M ${centerX} ${centerY} L ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 1 1 ${centerX - Math.cos(angle) * outerRadius} ${centerY - Math.sin(angle) * outerRadius} A ${outerRadius} ${outerRadius} 0 1 1 ${startX} ${startY} Z` : `M ${centerX} ${centerY} L ${startX} ${startY} A ${outerRadius} ${outerRadius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${endX} ${endY} Z`;
            slices.push({
              d,
              color: point.color
            });
            const label = formatDataLabel(series, point.index, point.value, point.text, total);
            if (label !== void 0) {
              const labelAngle = angle + sweep / 2;
              const position2 = labelPosition(series, point.index, "ctr");
              const labelRadius = doughnut ? position2 === "outEnd" ? outerRadius * 1.12 : position2 === "inEnd" ? innerRadius + ringWidth * 0.82 : position2 === "inBase" ? innerRadius + ringWidth * 0.18 : (outerRadius + innerRadius) / 2 : radius * (position2 === "outEnd" ? 1.12 : position2 === "inEnd" ? 0.82 : position2 === "inBase" ? 0.24 : position2 === "bestFit" ? 0.62 : 0.68);
              const x = position2 === "l" ? centerX - radius * 0.75 : position2 === "r" ? centerX + radius * 0.75 : centerX + Math.cos(labelAngle) * labelRadius;
              const y = position2 === "t" ? centerY - radius * 0.78 : position2 === "b" ? centerY + radius * 0.78 : centerY + Math.sin(labelAngle) * labelRadius + 6;
              pushDataLabel(label, x, y, position2 === "l" ? "end" : position2 === "r" ? "start" : "middle", series, point.index);
            }
            angle = endAngle;
          }
          if (seriesIndex === 0) points.slice(0, 12).forEach((point) => {
            var _a3, _b2;
            if ((_b2 = (_a3 = chart.legend) == null ? void 0 : _a3.hiddenEntries) == null ? void 0 : _b2.includes(point.index)) return;
            legend.push({
              text: point.text,
              x: 590,
              y: 105 + legend.length * 28,
              color: point.color
            });
          });
        });
      } else if (radar) {
        const radialValues = ticks.map((tick) => axisFraction(tick.value)).filter((fraction) => fraction > 1e-4 && fraction <= 1.0001);
        if (!((_ga = chart.valueAxis) == null ? void 0 : _ga.deleted) && ((_ha = chart.valueAxis) == null ? void 0 : _ha.majorGridlines)) for (const fraction of [.../* @__PURE__ */ new Set([...radialValues, 1])].sort((a, b) => a - b)) {
          const points = chart.categories.map((_category, index) => radarPoint(index, fraction));
          radarGrid.push({ d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z` });
        }
        if (!chart.categoryAxisDeleted && chart.categoryAxisMajorGridlines) chart.categories.forEach((_category, index) => {
          if (!showCategoryAxisTick(index)) return;
          const point = radarPoint(index);
          radarSpokes.push({
            x1: radarCenterX,
            y1: radarCenterY,
            x2: point.x,
            y2: point.y
          });
        });
        if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== "none") chart.categories.forEach((text, index) => {
          if (!showCategoryAxisLabel(index)) return;
          const point = radarPoint(index, 1, 22);
          const horizontal = Math.cos(point.angle);
          categories.push({
            text: formatCategoryAxisLabel(text),
            x: point.x,
            y: point.y + (Math.sin(point.angle) > 0.55 ? 16 : Math.sin(point.angle) < -0.55 ? -8 : 5),
            anchor: horizontal > 0.25 ? "start" : horizontal < -0.25 ? "end" : "middle"
          });
        });
        chart.series.forEach((series) => {
          var _a2, _b2;
          const points = chart.categories.map((_category, index) => {
            const value = series.values[index];
            return value == null ? null : radarPoint(index, axisFraction(value));
          });
          const complete = points.length >= 3 && points.every((point) => point !== null);
          if (chart.radarStyle === "filled" && complete) areas.push({
            d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z`,
            color: series.color,
            opacity: 0.22
          });
          if (series.radarShowLine !== false) {
            const path = complete ? `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z` : points.flatMap((point, index) => {
              const next = points[(index + 1) % points.length];
              return point && next ? [`M ${point.x} ${point.y} L ${next.x} ${next.y}`] : [];
            }).join(" ");
            if (path) lines.push({
              d: path,
              color: series.color
            });
          }
          points.forEach((point, index) => {
            var _a3, _b3, _c2, _d2, _e2, _f2;
            if (!point) return;
            const pointMarkerSymbol = (_b3 = (_a3 = series.markerSymbolOverrides) == null ? void 0 : _a3[index]) != null ? _b3 : series.markerSymbol;
            if (pointMarkerSymbol === void 0 ? series.radarShowMarker : pointMarkerSymbol !== "none") markers.push({
              x: point.x,
              y: point.y,
              color: series.color,
              symbol: pointMarkerSymbol != null ? pointMarkerSymbol : "circle",
              radius: (_f2 = (_e2 = (_d2 = (_c2 = series.markerSizeOverrides) == null ? void 0 : _c2[index]) != null ? _d2 : series.markerSize) != null ? _e2 : series.radarMarkerSize) != null ? _f2 : 5
            });
            const value = series.values[index];
            if (value == null) return;
            const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
            if (label === void 0) return;
            const position2 = labelPosition(series, index, "outEnd");
            const offset = position2 === "inBase" || position2 === "inEnd" || position2 === "ctr" ? -10 : 14;
            const x = position2 === "l" ? point.x - 12 : position2 === "r" ? point.x + 12 : point.x + Math.cos(point.angle) * offset;
            const y = position2 === "t" ? point.y - 14 : position2 === "b" ? point.y + 22 : point.y + Math.sin(point.angle) * offset + 6;
            const anchor = position2 === "l" ? "end" : position2 === "r" ? "start" : Math.cos(point.angle) > 0.25 ? "start" : Math.cos(point.angle) < -0.25 ? "end" : "middle";
            pushDataLabel(label, x, y, anchor, series, index);
          });
          if (!((_b2 = (_a2 = chart.legend) == null ? void 0 : _a2.hiddenEntries) == null ? void 0 : _b2.includes(series.legendIndex))) legend.push({
            text: series.name,
            x: 88 + legend.length * 220,
            y: 530,
            ...chartLegendKey(series)
          });
        });
      } else {
        const barSeriesByGroup = /* @__PURE__ */ new Map();
        chart.series.forEach((series, seriesIndex) => {
          var _a2;
          if (seriesType(series) !== "bar") return;
          const groupIndex = (_a2 = series.chartGroupIndex) != null ? _a2 : 0;
          barSeriesByGroup.set(groupIndex, [...barSeriesByGroup.get(groupIndex) || [], seriesIndex]);
        });
        for (let index = 0; index < categoryCount; index++) {
          const centerX = categoryX(index);
          const centerY = categoryY(index);
          if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== "none") {
            const categoryLabel = categoryTickLabelPosition === "high" ? horizontalValueAxis ? {
              x: ((_ia = chart.valueAxis) == null ? void 0 : _ia.reverse) ? plot.left - 10 : plot.left + plot.width + 10,
              y: centerY + 5,
              anchor: ((_ja = chart.valueAxis) == null ? void 0 : _ja.reverse) ? "end" : "start"
            } : {
              x: centerX,
              y: ((_ka = chart.valueAxis) == null ? void 0 : _ka.reverse) ? plot.top + plot.height + 25 : plot.top - 10,
              anchor: "middle"
            } : categoryTickLabelPosition === "low" ? horizontalValueAxis ? {
              x: ((_la = chart.valueAxis) == null ? void 0 : _la.reverse) ? plot.left + plot.width + 10 : plot.left - 10,
              y: centerY + 5,
              anchor: ((_ma = chart.valueAxis) == null ? void 0 : _ma.reverse) ? "start" : "end"
            } : {
              x: centerX,
              y: ((_na = chart.valueAxis) == null ? void 0 : _na.reverse) ? plot.top - 10 : plot.top + plot.height + 25,
              anchor: "middle"
            } : categoryAxisSide === "l" ? {
              x: categoryAxisCrossingCoordinate - 10,
              y: centerY + 5,
              anchor: "end"
            } : categoryAxisSide === "r" ? {
              x: categoryAxisCrossingCoordinate + 10,
              y: centerY + 5,
              anchor: "start"
            } : categoryAxisSide === "t" ? {
              x: centerX,
              y: categoryAxisCrossingCoordinate - 10,
              anchor: "middle"
            } : {
              x: centerX,
              y: categoryAxisCrossingCoordinate + 25,
              anchor: "middle"
            };
            if (showCategoryAxisLabel(index)) categories.push({
              text: formatCategoryAxisLabel(chart.categories[index] || ""),
              ...categoryLabel
            });
          }
          chart.series.forEach((series, seriesIndex) => {
            var _a2, _b2, _c2, _d2, _e2, _f2, _g2, _h2, _i2, _j2, _k2, _l2, _m2, _n2, _o2, _p2, _q2, _r2;
            const value = series.values[index];
            if (value === null || value === void 0) return;
            const color = series.color;
            const kind = seriesType(series);
            const stackedSegment = (_a2 = stackedSegmentFor(seriesIndex)) == null ? void 0 : _a2[index];
            const isStacked = (series.grouping || chart.grouping || "clustered") !== "clustered";
            if (kind === "line") {
              const x = horizontalValueAxis ? seriesValuePosition(series, (_b2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _b2 : value) : centerX;
              const y = horizontalValueAxis ? centerY : seriesValuePosition(series, (_c2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _c2 : value);
              const previousValue = index > 0 ? series.values[index - 1] : null;
              const previousStackedSegment = (_d2 = stackedSegmentFor(seriesIndex)) == null ? void 0 : _d2[index - 1];
              const previous = previousValue == null ? void 0 : horizontalValueAxis ? {
                x: seriesValuePosition(series, (_e2 = previousStackedSegment == null ? void 0 : previousStackedSegment.end) != null ? _e2 : previousValue),
                y: categoryY(index - 1)
              } : {
                x: categoryX(index - 1),
                y: seriesValuePosition(series, (_f2 = previousStackedSegment == null ? void 0 : previousStackedSegment.end) != null ? _f2 : previousValue)
              };
              if (series.lineShowLine !== false) lines.push({
                d: previous ? `M ${previous.x} ${previous.y} L ${x} ${y}` : `M ${x} ${y}`,
                color
              });
              const pointMarkerSymbol = (_h2 = (_g2 = series.markerSymbolOverrides) == null ? void 0 : _g2[index]) != null ? _h2 : series.markerSymbol;
              if (pointMarkerSymbol === void 0 ? series.lineShowMarker !== false : pointMarkerSymbol !== "none") markers.push({
                x,
                y,
                color,
                symbol: pointMarkerSymbol != null ? pointMarkerSymbol : "circle",
                radius: (_j2 = (_i2 = series.markerSizeOverrides) == null ? void 0 : _i2[index]) != null ? _j2 : series.markerSize
              });
              const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
              if (label !== void 0) {
                const position2 = labelPosition(series, index, "t");
                const labelX = position2 === "l" ? x - 10 : position2 === "r" ? x + 10 : x;
                const labelY = position2 === "b" ? y + 24 : position2 === "ctr" || position2 === "inBase" || position2 === "inEnd" ? y + 7 : y - 12;
                pushDataLabel(label, labelX, labelY, position2 === "l" ? "end" : position2 === "r" ? "start" : "middle", series, index);
              }
            } else if (kind === "area") {
              const x = horizontalValueAxis ? seriesValuePosition(series, (_k2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _k2 : value) : centerX;
              const y = horizontalValueAxis ? centerY : seriesValuePosition(series, (_l2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _l2 : value);
              const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
              if (label !== void 0) {
                const position2 = labelPosition(series, index, "t");
                const labelX = position2 === "l" ? x - 10 : position2 === "r" ? x + 10 : x;
                const labelY = position2 === "b" ? y + 24 : position2 === "ctr" || position2 === "inBase" || position2 === "inEnd" ? y + 7 : y - 12;
                pushDataLabel(label, labelX, labelY, position2 === "l" ? "end" : position2 === "r" ? "start" : "middle", series, index);
              }
            } else if (chart.direction === "horizontal") {
              const groupIndexes = barSeriesByGroup.get((_m2 = series.chartGroupIndex) != null ? _m2 : 0) || [seriesIndex];
              const barOrdinal = groupIndexes.indexOf(seriesIndex);
              const seriesCount = Math.max(groupIndexes.length, 1);
              const { minimum: axisMinimum, upper: axisUpper } = axisBoundsForSeries(series);
              const groupHeight = categoryBandSize(index) * 0.72;
              const barHeight = groupHeight / (isStacked ? 1 : seriesCount);
              const barY = centerY - groupHeight / 2 + (isStacked ? 0 : barOrdinal * barHeight);
              const zeroX = seriesValuePosition(series, Math.max(axisMinimum, Math.min(axisUpper, 0)));
              const baseX = isStacked ? seriesValuePosition(series, (_n2 = stackedSegment == null ? void 0 : stackedSegment.start) != null ? _n2 : 0) : zeroX;
              const valueX = seriesValuePosition(series, isStacked ? (_o2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _o2 : value : value);
              bars.push({
                x: Math.min(baseX, valueX),
                y: barY,
                width: Math.max(1, Math.abs(valueX - baseX)),
                height: barHeight * 0.86,
                color
              });
              const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
              if (label !== void 0) {
                const position2 = labelPosition(series, index, isStacked ? "ctr" : "outEnd");
                const direction = value >= 0 ? 1 : -1;
                const left = Math.min(baseX, valueX);
                const right = Math.max(baseX, valueX);
                let labelX = valueX + direction * 8;
                let labelY = barY + barHeight * 0.68;
                let anchor = direction > 0 ? "start" : "end";
                if (position2 === "ctr") {
                  labelX = (baseX + valueX) / 2;
                  anchor = "middle";
                } else if (position2 === "inBase") {
                  labelX = baseX + direction * 20;
                  anchor = direction > 0 ? "start" : "end";
                } else if (position2 === "inEnd") {
                  labelX = valueX - direction * 20;
                  anchor = direction > 0 ? "end" : "start";
                } else if (position2 === "l") {
                  labelX = left - 8;
                  anchor = "end";
                } else if (position2 === "r") {
                  labelX = right + 8;
                  anchor = "start";
                } else if (position2 === "t") {
                  labelX = (zeroX + valueX) / 2;
                  labelY = barY - 8;
                  anchor = "middle";
                } else if (position2 === "b") {
                  labelX = (zeroX + valueX) / 2;
                  labelY = barY + barHeight + 15;
                  anchor = "middle";
                }
                pushDataLabel(label, labelX, labelY, anchor, series, index);
              }
            } else {
              const groupIndexes = barSeriesByGroup.get((_p2 = series.chartGroupIndex) != null ? _p2 : 0) || [seriesIndex];
              const barOrdinal = groupIndexes.indexOf(seriesIndex);
              const seriesCount = Math.max(groupIndexes.length, 1);
              const { minimum: axisMinimum, upper: axisUpper } = axisBoundsForSeries(series);
              const groupWidth = categoryBandSize(index) * 0.72;
              const barWidth = groupWidth / (isStacked ? 1 : seriesCount);
              const zeroY = seriesValuePosition(series, Math.max(axisMinimum, Math.min(axisUpper, 0)));
              const baseY = isStacked ? seriesValuePosition(series, (_q2 = stackedSegment == null ? void 0 : stackedSegment.start) != null ? _q2 : 0) : zeroY;
              const valueY = seriesValuePosition(series, isStacked ? (_r2 = stackedSegment == null ? void 0 : stackedSegment.end) != null ? _r2 : value : value);
              const barX = centerX - groupWidth / 2 + (isStacked ? 0 : barOrdinal * barWidth);
              bars.push({
                x: barX,
                y: Math.min(baseY, valueY),
                width: barWidth * 0.86,
                height: Math.max(1, Math.abs(valueY - baseY)),
                color
              });
              const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
              if (label !== void 0) {
                const position2 = labelPosition(series, index, isStacked ? "ctr" : "outEnd");
                const top = Math.min(baseY, valueY);
                const bottom = Math.max(baseY, valueY);
                let labelX = barX + barWidth * 0.43;
                let labelY = value >= 0 ? top - 8 : bottom + 20;
                let anchor = "middle";
                if (position2 === "ctr") labelY = (baseY + valueY) / 2 + 6;
                else if (position2 === "inBase") labelY = value >= 0 ? baseY - 8 : baseY + 18;
                else if (position2 === "inEnd") labelY = value >= 0 ? valueY + 18 : valueY - 4;
                else if (position2 === "b") labelY = bottom + 20;
                else if (position2 === "l") {
                  labelX = barX - 8;
                  labelY = (baseY + valueY) / 2 + 6;
                  anchor = "end";
                } else if (position2 === "r") {
                  labelX = barX + barWidth * 0.86 + 8;
                  labelY = (baseY + valueY) / 2 + 6;
                  anchor = "start";
                }
                pushDataLabel(label, labelX, labelY, anchor, series, index);
              }
            }
          });
        }
        chart.series.slice(0, 12).forEach((series) => {
          var _a2, _b2;
          if ((_b2 = (_a2 = chart.legend) == null ? void 0 : _a2.hiddenEntries) == null ? void 0 : _b2.includes(series.legendIndex)) return;
          legend.push({
            text: series.name,
            x: 88 + legend.length % 4 * 220,
            y: 530 + Math.floor(legend.length / 4) * 22,
            ...chartLegendKey(series)
          });
        });
      }
      if (!chart.legend) legend.length = 0;
      else {
        const pos = chart.legend.position;
        const defaultX = defaultLegendOrigin(pos, "x");
        const defaultY = defaultLegendOrigin(pos, "y");
        const manual = chart.legend.manualLayout;
        const originX = (manual == null ? void 0 : manual.x) ? manual.x.mode === "edge" ? manual.x.value * 1e3 : defaultX + manual.x.value * 1e3 : defaultX;
        const originY = (manual == null ? void 0 : manual.y) ? manual.y.mode === "edge" ? manual.y.value * 600 : defaultY + manual.y.value * 600 : defaultY;
        const offsetX = originX - defaultX;
        const offsetY = originY - defaultY;
        legend.forEach((entry, index) => {
          const x = pos === "left" ? 88 : pos === "right" ? 780 : pos === "topRight" ? 760 : 88 + index % 4 * 220;
          const y = pos === "left" || pos === "right" ? 105 + index * 28 : pos === "top" ? 77 + Math.floor(index / 4) * 18 : pos === "topRight" ? 92 + index * 22 : 530 + Math.floor(index / 4) * 22;
          legend[index] = {
            ...entry,
            x: x + offsetX,
            y: y + offsetY
          };
        });
      }
      const manualLayout = (_oa = chart.legend) == null ? void 0 : _oa.manualLayout;
      let legendClip;
      if ((manualLayout == null ? void 0 : manualLayout.w) || (manualLayout == null ? void 0 : manualLayout.h)) {
        const x = ((_pa = manualLayout.x) == null ? void 0 : _pa.mode) === "edge" ? manualLayout.x.value * 1e3 : defaultLegendOrigin((_qa = chart.legend) == null ? void 0 : _qa.position, "x") + (((_ra = manualLayout.x) == null ? void 0 : _ra.value) || 0) * 1e3;
        const y = ((_sa = manualLayout.y) == null ? void 0 : _sa.mode) === "edge" ? manualLayout.y.value * 600 : defaultLegendOrigin((_ta = chart.legend) == null ? void 0 : _ta.position, "y") + (((_ua = manualLayout.y) == null ? void 0 : _ua.value) || 0) * 600;
        const width = ((_va = manualLayout.w) == null ? void 0 : _va.mode) === "edge" ? manualLayout.w.value * 1e3 - x : ((_xa = (_wa = manualLayout.w) == null ? void 0 : _wa.value) != null ? _xa : 1) * 1e3;
        const height = ((_ya = manualLayout.h) == null ? void 0 : _ya.mode) === "edge" ? manualLayout.h.value * 600 - y : ((_Aa = (_za = manualLayout.h) == null ? void 0 : _za.value) != null ? _Aa : 1) * 600;
        legendClip = {
          id: `${customGradientScope}-legend-${elementId}`.replace(/[^a-zA-Z0-9_-]/g, "_"),
          x,
          y,
          width: Math.max(0, width),
          height: Math.max(0, height)
        };
      }
      const baseline = xyChart ? {
        x1: plot.left,
        y1: xyXAxisCrossingY,
        x2: plot.left + plot.width,
        y2: xyXAxisCrossingY
      } : horizontalValueAxis ? {
        x1: categoryAxisCrossingCoordinate,
        y1: plot.top,
        x2: categoryAxisCrossingCoordinate,
        y2: plot.top + plot.height
      } : {
        x1: plot.left,
        y1: categoryAxisCrossingCoordinate,
        x2: plot.left + plot.width,
        y2: categoryAxisCrossingCoordinate
      };
      if ((dateAxisMajorTickValues == null ? void 0 : dateAxisMajorTickValues.length) && categories.length && hasDateAxisScale) {
        const firstLabel = categories[0];
        categories.splice(0, categories.length, ...dateAxisMajorTickValues.map((value) => {
          const fraction = dateAxisFraction(value);
          return {
            ...firstLabel,
            text: formatCategoryAxisLabel(String(value)),
            ...horizontalValueAxis ? { y: plot.top + plot.height * fraction } : { x: plot.left + plot.width * fraction }
          };
        }));
      }
      return {
        plot,
        ticks,
        xTicks,
        axisTicks,
        secondaryValueAxes,
        radarGrid,
        radarSpokes,
        categoryGridlines,
        categories,
        bars,
        areas,
        lines,
        markers,
        slices,
        legend,
        ...legendClip ? { legendClip } : {},
        dataLabels,
        baseline,
        valueAxisLine: radar ? void 0 : valueAxisLine
      };
    }
    const props = __props;
    const emit = __emit;
    const backgroundImageSize = ref();
    const backgroundTileWarning = ref("");
    watch(() => {
      var _a, _b, _c, _d, _e, _f, _g, _h;
      return [
        (_a = props.slide.backgroundImage) == null ? void 0 : _a.url,
        Boolean((_b = props.slide.backgroundImage) == null ? void 0 : _b.tile),
        ((_d = (_c = props.slide.backgroundImage) == null ? void 0 : _c.tile) == null ? void 0 : _d.flip) || "none",
        (_e = props.slide.backgroundImage) == null ? void 0 : _e.crop.left,
        (_f = props.slide.backgroundImage) == null ? void 0 : _f.crop.top,
        (_g = props.slide.backgroundImage) == null ? void 0 : _g.crop.right,
        (_h = props.slide.backgroundImage) == null ? void 0 : _h.crop.bottom
      ];
    }, ([url, tiled, flip], _oldValue, onCleanup) => {
      backgroundImageSize.value = void 0;
      backgroundTileWarning.value = "";
    }, { immediate: true });
    const frameStyle = computed(() => {
      var _a, _b, _c, _d, _e;
      const image = props.slide.backgroundImage;
      if (!image) return {
        aspectRatio: `${props.width} / ${props.height}`,
        background: props.slide.background
      };
      const { left, top, right, bottom } = image.crop;
      const tile = image.tile;
      const alignments = {
        tl: [0, 0],
        t: [50, 0],
        tr: [100, 0],
        l: [0, 50],
        ctr: [50, 50],
        r: [100, 50],
        bl: [0, 100],
        b: [50, 100],
        br: [100, 100]
      };
      const naturalSize = ((_a = backgroundImageSize.value) == null ? void 0 : _a.sourceUrl) === image.url ? backgroundImageSize.value : void 0;
      const tileWidth = naturalSize && tile ? naturalSize.width * tile.scaleX : 0;
      const tileHeight = naturalSize && tile ? naturalSize.height * tile.scaleY : 0;
      const tilePosition = (alignment, offset) => offset === 0 ? `${alignment}%` : `calc(${alignment}% ${offset < 0 ? "-" : "+"} ${Math.abs(offset)}px)`;
      const position = tile && tile.align === "tl" && tile.offsetX === 0 && tile.offsetY === 0 && ((_b = naturalSize == null ? void 0 : naturalSize.repeatX) != null ? _b : 1) === 1 && ((_c = naturalSize == null ? void 0 : naturalSize.repeatY) != null ? _c : 1) === 1 ? "0px 0px" : tile ? `${tilePosition(alignments[tile.align][0], tile.offsetX * scale.value + tileWidth * alignments[tile.align][0] / 100 * (((_d = naturalSize == null ? void 0 : naturalSize.repeatX) != null ? _d : 1) - 1))} ${tilePosition(alignments[tile.align][1], tile.offsetY * scale.value + tileHeight * alignments[tile.align][1] / 100 * (((_e = naturalSize == null ? void 0 : naturalSize.repeatY) != null ? _e : 1) - 1))}` : `${left + right ? left / (left + right) * 100 : 50}% ${top + bottom ? top / (top + bottom) * 100 : 50}%`;
      return {
        aspectRatio: `${props.width} / ${props.height}`,
        background: props.slide.background,
        backgroundImage: `url("${(naturalSize == null ? void 0 : naturalSize.url) || image.url}")`,
        backgroundRepeat: tile ? "repeat" : "no-repeat",
        backgroundSize: tile ? naturalSize ? `${tileWidth * naturalSize.repeatX}px ${tileHeight * naturalSize.repeatY}px` : "auto" : `${100 / (1 - (left + right) / 1e5)}% ${100 / (1 - (top + bottom) / 1e5)}%`,
        backgroundPosition: position
      };
    });
    const stage = ref(null);
    const stageWidth = ref(0);
    const autoFitDimensions = reactive(/* @__PURE__ */ new Map());
    const normalAutoFitScales = reactive(/* @__PURE__ */ new Map());
    const textEllipsisLineCounts = reactive(/* @__PURE__ */ new Map());
    const elementLayoutVersion = ref(0);
    const textOverflowWarning = ref("");
    const animationClock = ref(0);
    const reduceMotion = ref(false);
    const invalidMediaTrimIds = ref([]);
    const invalidMediaFadeIds = ref([]);
    const blockedMediaPlaybackIds = ref([]);
    const lastMediaBookmarkTimes = /* @__PURE__ */ new WeakMap();
    const mediaBookmarkPlaybackStarted = /* @__PURE__ */ new WeakSet();
    computed(() => {
      var _a;
      const triggersByMedia = /* @__PURE__ */ new Map();
      (_a = props.slide.triggeredAnimations) == null ? void 0 : _a.forEach((sequence, sequenceIndex) => {
        const trigger = sequence.trigger;
        if (trigger.type !== "mediaBookmark") return;
        const triggers = triggersByMedia.get(trigger.id) || [];
        triggers.push({
          sequenceIndex,
          timeMs: trigger.timeMs
        });
        triggersByMedia.set(trigger.id, triggers);
      });
      return triggersByMedia;
    });
    const executedMediaActions = /* @__PURE__ */ new Set();
    const completedMediaActions = reactive(/* @__PURE__ */ new Map());
    const mediaActionEndTimers = /* @__PURE__ */ new Map();
    let automaticStartedAt;
    let stepStartedAt = /* @__PURE__ */ new Map();
    let mainStepEndedAt = /* @__PURE__ */ new Map();
    let triggerSteps = /* @__PURE__ */ new Map();
    let triggerStartedAt = /* @__PURE__ */ new Map();
    let triggerStepEndedAt = /* @__PURE__ */ new Map();
    let triggerPlaybackByKey = /* @__PURE__ */ new Map();
    const firedTimingNodeTriggers = /* @__PURE__ */ new Set();
    let autoFitFrame = 0;
    let animationFrame = 0;
    let animationTimer;
    const pixelRatio = ref(1);
    const scale = computed(() => stageWidth.value / Math.max(props.width, 1));
    const chartVisuals = computed(() => new Map(props.slide.elements.flatMap((element) => element.chart ? [[element.id, makeChartVisual(element.chart, element.id)]] : [])));
    const chartDataLabelKeyPositions = ref(/* @__PURE__ */ new Map());
    const chartDataLabelKeyId = (elementId, index) => `${customGradientScope}:${elementId}:${index}`;
    const refreshChartDataLabelKeyPositions = async () => {
      await nextTick();
    };
    watch(chartVisuals, () => {
      refreshChartDataLabelKeyPositions();
    }, { flush: "post" });
    function supportsTextVerticalEllipsis(element) {
      var _a;
      return element.textVerticalOverflow === "ellipsis" && ((_a = element.textColumnCount) != null ? _a : 1) === 1 && (!element.textOrientation || element.textOrientation === "horz");
    }
    function measureTextEllipsisLineCount(frame) {
      const frameRect = frame.getBoundingClientRect();
      const paddingBottom = Number.parseFloat(getComputedStyle(frame).paddingBottom) || 0;
      const contentBottom = frameRect.bottom - paddingBottom;
      const bands = [];
      for (const run of frame.querySelectorAll(".text-run")) {
        const range = (void 0).createRange();
        range.selectNodeContents(run);
        for (const rect of range.getClientRects()) {
          if (rect.width <= 0 || rect.height <= 0) continue;
          const band = bands.find((line) => rect.top < line.bottom - 1 && rect.bottom > line.top + 1);
          if (band) {
            band.top = Math.min(band.top, rect.top);
            band.bottom = Math.max(band.bottom, rect.bottom);
          } else bands.push({
            top: rect.top,
            bottom: rect.bottom
          });
        }
      }
      bands.sort((left, right) => left.top - right.top);
      const visibleLines = bands.filter((line) => line.bottom <= contentBottom + 0.5).length;
      return bands.some((line) => line.bottom > contentBottom + 0.5) ? Math.max(1, visibleLines) : 0;
    }
    function scheduleShapeAutoFit() {
      if (!stage.value || autoFitFrame || scale.value <= 0) return;
      autoFitFrame = requestAnimationFrame(() => {
        var _a, _b, _c, _d, _e, _f, _g, _h;
        autoFitFrame = 0;
        if (!stage.value) return;
        const ellipsisFrames = [...stage.value.querySelectorAll('.text-frame[data-text-overflow="ellipsis"]')];
        let resetTextEllipsis = false;
        let layoutChanged = false;
        for (const frame of ellipsisFrames) {
          const id = frame.dataset.elementId;
          if (!id) continue;
          const deleted = textEllipsisLineCounts.delete(`${props.slide.id}:${id}`);
          layoutChanged || (layoutChanged = deleted);
          resetTextEllipsis = deleted || resetTextEllipsis;
        }
        if (resetTextEllipsis) {
          elementLayoutVersion.value++;
          nextTick(scheduleShapeAutoFit);
          return;
        }
        const canClampText = typeof CSS !== "undefined" && CSS.supports("display", "-webkit-box") && CSS.supports("-webkit-line-clamp", "1");
        textOverflowWarning.value = ellipsisFrames.length && !canClampText ? "Vertical text ellipsis is clipped because this browser does not support CSS line clamping." : "";
        if (canClampText) for (const frame of ellipsisFrames) {
          const id = frame.dataset.elementId;
          if (!id) continue;
          const lineCount = measureTextEllipsisLineCount(frame);
          if (lineCount > 0 && textEllipsisLineCounts.get(`${props.slide.id}:${id}`) !== lineCount) {
            textEllipsisLineCounts.set(`${props.slide.id}:${id}`, lineCount);
            layoutChanged = true;
          }
        }
        const elements = new Map(props.slide.elements.map((element) => [element.id, element]));
        let changed = false;
        for (const frame of stage.value.querySelectorAll('.text-frame[data-autofit="shape"]')) {
          const id = frame.dataset.elementId;
          const element = id ? elements.get(id) : void 0;
          if (!id || !element) continue;
          const overflowWidth = Math.max(0, frame.scrollWidth - frame.clientWidth);
          const overflowHeight = Math.max(0, frame.scrollHeight - frame.clientHeight);
          if (!overflowWidth && !overflowHeight) continue;
          const key = `${props.slide.id}:${id}`;
          const current = autoFitDimensions.get(key) || {
            width: element.width,
            height: element.height
          };
          autoFitDimensions.set(key, {
            width: current.width + overflowWidth / scale.value,
            height: current.height + overflowHeight / scale.value
          });
          changed = true;
          layoutChanged = true;
        }
        for (const frame of stage.value.querySelectorAll('.text-frame[data-autofit="normal"]')) {
          const id = frame.dataset.elementId;
          const element = id ? elements.get(id) : void 0;
          if (!id || !element || !frame.clientWidth || !frame.clientHeight) continue;
          const overflowWidth = frame.scrollWidth > frame.clientWidth + 1 ? frame.clientWidth / frame.scrollWidth : 1;
          const overflowHeight = frame.scrollHeight > frame.clientHeight + 1 ? frame.clientHeight / frame.scrollHeight : 1;
          if (overflowWidth === 1 && overflowHeight === 1) continue;
          const key = `${props.slide.id}:${id}`;
          const current = (_b = (_a = normalAutoFitScales.get(key)) != null ? _a : element.textFontScale) != null ? _b : 1;
          const fitted = Math.max(0.05, current * Math.min(overflowWidth, overflowHeight) * 0.98);
          if (fitted < current - 1e-3) {
            normalAutoFitScales.set(key, fitted);
            changed = true;
            layoutChanged = true;
          }
        }
        for (const cell of stage.value.querySelectorAll('.slide-table td[data-autofit="normal"]')) {
          const id = (_c = cell.closest(".slide-element")) == null ? void 0 : _c.dataset.elementId;
          const rowIndex = Number(cell.dataset.rowIndex);
          const cellIndex = Number(cell.dataset.cellIndex);
          const model = (_f = (_e = (_d = id ? elements.get(id) : void 0) == null ? void 0 : _d.table) == null ? void 0 : _e.rows[rowIndex]) == null ? void 0 : _f.cells[cellIndex];
          const text = cell.querySelector(".table-cell-text");
          if (!id || !model || !(text == null ? void 0 : text.clientWidth) || !text.clientHeight) continue;
          const overflowWidth = text.scrollWidth > text.clientWidth + 1 ? text.clientWidth / text.scrollWidth : 1;
          const overflowHeight = text.scrollHeight > text.clientHeight + 1 ? text.clientHeight / text.scrollHeight : 1;
          if (overflowWidth === 1 && overflowHeight === 1) continue;
          const key = tableCellAutoFitKey(id, rowIndex, cellIndex);
          const current = (_h = (_g = normalAutoFitScales.get(key)) != null ? _g : model.textFontScale) != null ? _h : 1;
          const fitted = Math.max(0.05, current * Math.min(overflowWidth, overflowHeight) * 0.98);
          if (fitted < current - 1e-3) {
            normalAutoFitScales.set(key, fitted);
            changed = true;
            layoutChanged = true;
          }
        }
        if (layoutChanged) elementLayoutVersion.value++;
        if (changed) nextTick(scheduleShapeAutoFit);
      });
    }
    const chartPlot = (id) => {
      var _a, _b;
      return (_b = (_a = chartVisuals.value.get(id)) == null ? void 0 : _a.plot) != null ? _b : {
        left: 88,
        top: 75,
        width: 850,
        height: 390
      };
    };
    const triggerKey = (sequence, step) => `${sequence}:${step}`;
    const timelineEnd = (action) => action.triggerSequence !== void 0 && action.triggerStep !== void 0 ? triggerStepEndedAt.get(triggerKey(action.triggerSequence, action.triggerStep)) : action.automatic ? void 0 : mainStepEndedAt.get(action.step);
    const timelineNow = (action) => {
      var _a;
      return Math.min(animationClock.value, (_a = timelineEnd(action)) != null ? _a : Infinity);
    };
    const timelineStart = (action) => {
      var _a;
      const startedAt = action.triggerSequence === void 0 || action.triggerStep === void 0 ? action.automatic ? automaticStartedAt : stepStartedAt.get(action.step) : triggerStartedAt.get(triggerKey(action.triggerSequence, action.triggerStep));
      if (startedAt === void 0 || startedAt === Number.NEGATIVE_INFINITY || !((_a = action.mediaWaitForEndKeys) == null ? void 0 : _a.length)) return startedAt;
      const endedAt = action.mediaWaitForEndKeys.map((key) => completedMediaActions.get(key));
      if (endedAt.some((time) => time === void 0)) return Infinity;
      return Math.max(startedAt, ...endedAt.map((time) => time + (action.mediaWaitDelayMs || 0) - action.delayMs));
    };
    const animationDuration = (action) => {
      var _a, _b;
      return action.durationMs === 0 ? 0 : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * ((_a = action.repeatCount) != null ? _a : 1), (_b = action.repeatDurationMs) != null ? _b : Infinity);
    };
    const naturalAnimationDuration = (action) => {
      var _a, _b, _c;
      return action.effect === "media" ? (_a = action.mediaDurationMs) != null ? _a : 0 : action.timingWarp ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs) : action.durationMs === 0 ? 0 : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (Number.isFinite(action.repeatCount) || action.repeatDurationMs !== void 0 ? (_b = action.repeatCount) != null ? _b : 1 : 1), (_c = action.repeatDurationMs) != null ? _c : Infinity) + (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0);
    };
    const animationScheduleDuration = (action) => action.effect === "media" ? 0 : action.timingWarp ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs) : animationDuration(action) + (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0);
    function iterationIndexFor(action, start, end) {
      var _a, _b;
      const ranges = (_a = action.iteration) == null ? void 0 : _a.ranges;
      if (!ranges) return void 0;
      let low = 0, high = ranges.length;
      while (low < high) {
        const middle = low + high >>> 1;
        if (ranges[middle].end <= start) low = middle + 1;
        else high = middle;
      }
      const range = ranges[low];
      return range && range.start <= start && range.end >= end ? ((_b = action.iteration) == null ? void 0 : _b.backwards) ? ranges.length - low - 1 : low : -1;
    }
    const animationsByTarget = computed(() => {
      var _a, _b;
      const targets = /* @__PURE__ */ new Map();
      for (const animation of props.slide.automaticAnimations || []) {
        const actions = targets.get(animation.targetId) || [];
        actions.push({
          ...animation,
          step: -1,
          automatic: true
        });
        targets.set(animation.targetId, actions);
      }
      (_a = props.slide.animationSteps) == null ? void 0 : _a.forEach((step, index) => {
        for (const animation of step) {
          const actions = targets.get(animation.targetId) || [];
          actions.push({
            ...animation,
            step: index,
            automatic: false
          });
          targets.set(animation.targetId, actions);
        }
      });
      (_b = props.slide.triggeredAnimations) == null ? void 0 : _b.forEach((sequence, sequenceIndex) => {
        sequence.steps.forEach((step, stepIndex) => {
          for (const animation of step) {
            const actions = targets.get(animation.targetId) || [];
            actions.push({
              ...animation,
              step: -1,
              automatic: false,
              triggerSequence: sequenceIndex,
              triggerStep: stepIndex
            });
            targets.set(animation.targetId, actions);
          }
        });
      });
      for (const actions of targets.values()) actions.sort((a, b) => Number(a.triggerSequence !== void 0) - Number(b.triggerSequence !== void 0) || a.step - b.step || a.delayMs - b.delayMs);
      return targets;
    });
    const hasInteractiveMedia = computed(() => !props.thumbnail && props.slide.elements.some((element) => element.kind === "video" || element.kind === "audio"));
    const timelineActions = computed(() => [
      ...(props.slide.automaticAnimations || []).map((animation) => ({
        ...animation,
        step: -1,
        automatic: true
      })),
      ...(props.slide.animationSteps || []).flatMap((step, index) => step.map((animation) => ({
        ...animation,
        step: index,
        automatic: false
      }))),
      ...(props.slide.triggeredAnimations || []).flatMap((sequence, sequenceIndex) => sequence.steps.flatMap((step, stepIndex) => step.map((animation) => ({
        ...animation,
        step: -1,
        automatic: false,
        triggerSequence: sequenceIndex,
        triggerStep: stepIndex
      }))))
    ]);
    function endTimelineStep(triggerSequence, step, endedAt) {
      if (triggerSequence === void 0) mainStepEndedAt.set(step, endedAt);
      else triggerStepEndedAt.set(triggerKey(triggerSequence, step), endedAt);
      timelineActions.value.forEach((action, index) => {
        var _a;
        if (!(triggerSequence === void 0 ? action.triggerSequence === void 0 && !action.automatic && action.step === step : action.triggerSequence === triggerSequence && action.triggerStep === step) || action.effect !== "media" || !executedMediaActions.has(index)) return;
        const media = Array.from(((_a = stage.value) == null ? void 0 : _a.querySelectorAll("[data-media-id]")) || []).find((item) => item.dataset.mediaId === action.targetId);
        if (media) {
          media.pause();
          delete media.dataset.mediaActionEndAt;
          mediaBookmarkPlaybackStarted.delete(media);
          lastMediaBookmarkTimes.delete(media);
        }
        if (action.mediaActionKey) {
          completedMediaActions.set(action.mediaActionKey, endedAt);
          const timer = mediaActionEndTimers.get(action.mediaActionKey);
          if (timer) clearTimeout(timer);
          mediaActionEndTimers.delete(action.mediaActionKey);
        }
      });
    }
    const mediaSlideCounts = computed(() => {
      const counts = /* @__PURE__ */ new Map();
      for (const action of timelineActions.value) if (action.effect === "media" && action.mediaSlideCount) counts.set(action.targetId, Math.max(action.mediaSlideCount, counts.get(action.targetId) || 0));
      return counts;
    });
    function cancelTimelineUpdates() {
      cancelAnimationFrame(animationFrame);
      animationFrame = 0;
      if (animationTimer) clearTimeout(animationTimer);
      animationTimer = void 0;
    }
    function scheduleTimelineUpdate() {
      cancelTimelineUpdates();
      if (props.thumbnail && !props.playbackPreview) return;
      const nextTimingNodeTriggerAt = runTimingNodeTriggeredAnimations(animationClock.value);
      runTimedMediaActions();
      const now = animationClock.value;
      if (reduceMotion.value) {
        const nextScheduledAt = timelineActions.value.reduce((next, action) => {
          if (action.effect !== "media") return next;
          const startedAt = timelineStart(action);
          return startedAt === void 0 || startedAt === Number.NEGATIVE_INFINITY ? next : Math.min(next, startedAt + action.delayMs);
        }, nextTimingNodeTriggerAt);
        if (Number.isFinite(nextScheduledAt) && nextScheduledAt > now) animationTimer = setTimeout(() => {
          animationTimer = void 0;
          animationClock.value = performance.now();
          scheduleTimelineUpdate();
        }, Math.max(0, nextScheduledAt - now));
        return;
      }
      let nextStart = nextTimingNodeTriggerAt;
      let hasActiveAnimation = false;
      for (const action of timelineActions.value) {
        const startedAt = timelineStart(action);
        if (startedAt === void 0 || startedAt === Number.NEGATIVE_INFINITY) continue;
        const stoppedAt = timelineEnd(action);
        if (stoppedAt !== void 0 && stoppedAt <= now) continue;
        const start = startedAt + action.delayMs;
        const end = Math.min(start + animationScheduleDuration(action), stoppedAt != null ? stoppedAt : Infinity);
        if (start > now) nextStart = Math.min(nextStart, start);
        else if (end > now) hasActiveAnimation = true;
      }
      const update = () => {
        animationFrame = 0;
        animationTimer = void 0;
        animationClock.value = performance.now();
        scheduleTimelineUpdate();
      };
      if (hasActiveAnimation) animationFrame = requestAnimationFrame(update);
      else if (Number.isFinite(nextStart)) animationTimer = setTimeout(update, Math.max(0, nextStart - now));
    }
    function refreshTimeline() {
      cancelTimelineUpdates();
      animationClock.value = performance.now();
      scheduleTimelineUpdate();
    }
    function beginSlideTimeline(step) {
      const now = performance.now();
      invalidMediaTrimIds.value = [];
      invalidMediaFadeIds.value = [];
      blockedMediaPlaybackIds.value = [];
      executedMediaActions.clear();
      completedMediaActions.clear();
      for (const timer of mediaActionEndTimers.values()) clearTimeout(timer);
      mediaActionEndTimers.clear();
      automaticStartedAt = now;
      stepStartedAt = new Map(Array.from({ length: step }, (_, index) => [index, Number.NEGATIVE_INFINITY]));
      mainStepEndedAt.clear();
      triggerSteps.clear();
      triggerStartedAt.clear();
      triggerStepEndedAt.clear();
      triggerPlaybackByKey.clear();
      firedTimingNodeTriggers.clear();
      animationClock.value = now;
      scheduleTimelineUpdate();
    }
    function syncTriggerPlayback() {
      var _a;
      if (props.triggerPlayback === void 0) return;
      triggerSteps.clear();
      triggerStartedAt.clear();
      triggerStepEndedAt.clear();
      triggerPlaybackByKey.clear();
      for (const playback of [...props.triggerPlayback].sort((a, b) => a.startedAt - b.startedAt)) {
        const sequence = (_a = props.slide.triggeredAnimations) == null ? void 0 : _a[playback.sequence];
        if (!sequence || !Number.isInteger(playback.step) || playback.step < 0 || playback.step >= sequence.steps.length || !Number.isFinite(playback.startedAt)) continue;
        const key = triggerKey(playback.sequence, playback.step);
        const startedAt = playback.startedAt - performance.timeOrigin;
        if (!sequence.concurrent && playback.step > 0) {
          const previousKey = triggerKey(playback.sequence, playback.step - 1);
          if (triggerStartedAt.has(previousKey)) endTimelineStep(playback.sequence, playback.step - 1, startedAt);
        }
        triggerStartedAt.set(key, startedAt);
        triggerPlaybackByKey.set(key, playback);
        triggerSteps.set(playback.sequence, Math.max(triggerSteps.get(playback.sequence) || 0, playback.step + 1));
      }
      refreshTimeline();
    }
    watch([() => props.slide.id, () => props.animationStep], ([slideId, step], [oldSlideId, oldStep]) => {
      var _a, _b, _c, _d, _e;
      if (!stage.value) return;
      if (slideId !== oldSlideId) {
        autoFitDimensions.clear();
        normalAutoFitScales.clear();
        textEllipsisLineCounts.clear();
        beginSlideTimeline(step);
        nextTick(scheduleShapeAutoFit);
        return;
      }
      if (step > oldStep) {
        for (let previous = oldStep + 1; previous < step; previous++) stepStartedAt.set(previous, Number.NEGATIVE_INFINITY);
        if (step === oldStep + 1) runSlideTriggeredAnimations("onNext");
        const now = performance.now();
        const previousStep = oldStep - 1;
        if (previousStep >= 0 && !((_a = props.slide.animationSequence) == null ? void 0 : _a.concurrent) && ((_b = props.slide.animationSequence) == null ? void 0 : _b.nextAction) !== "seek") endTimelineStep(void 0, previousStep, now);
        if (((_c = props.slide.animationSequence) == null ? void 0 : _c.nextAction) === "seek" && oldStep > 0) {
          const startedAt = stepStartedAt.get(previousStep);
          const naturalEnd = Math.max(0, ...(((_d = props.slide.animationSteps) == null ? void 0 : _d[previousStep]) || []).map((action) => action.delayMs + naturalAnimationDuration(action)));
          if (startedAt !== void 0 && startedAt !== Number.NEGATIVE_INFINITY) stepStartedAt.set(previousStep, Math.min(startedAt, now - naturalEnd));
        }
        stepStartedAt.set(oldStep, now);
      } else if (step < oldStep) {
        if (step === oldStep - 1) runSlideTriggeredAnimations("onPrev");
        (_e = props.slide.triggeredAnimations) == null ? void 0 : _e.forEach((sequence, index) => {
          if (sequence.trigger.type === "timeNode" && sequence.trigger.sourceStep !== void 0 && sequence.trigger.sourceStep >= step) firedTimingNodeTriggers.delete(index);
        });
        for (const startedStep of stepStartedAt.keys()) if (startedStep >= step) {
          stepStartedAt.delete(startedStep);
          mainStepEndedAt.delete(startedStep);
        }
        timelineActions.value.forEach((action, index) => {
          if (action.triggerSequence === void 0 && !action.automatic && action.step >= step) executedMediaActions.delete(index);
        });
      }
      refreshTimeline();
    }, { flush: "post" });
    watch([() => props.slide.id, () => props.triggerPlayback], syncTriggerPlayback, { flush: "post" });
    function applyColorOpacity(color, opacity, brightness = 0) {
      if (opacity >= 1 && !brightness || color === "transparent") return color;
      if (/(?:repeating-)?(?:linear|radial|conic)-gradient\(/i.test(color)) return color.replace(/#[\da-f]{6}\b|rgba?\([^)]*\)|\btransparent\b/gi, (stopColor) => applyColorOpacity(stopColor, opacity, brightness));
      const hex = /^#([\da-f]{6})$/i.exec(color);
      const rgb = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)(?:\s*,\s*([\d.]+)\s*)?\)$/i.exec(color);
      const channels = hex ? [
        parseInt(hex[1].slice(0, 2), 16),
        parseInt(hex[1].slice(2, 4), 16),
        parseInt(hex[1].slice(4, 6), 16),
        1
      ] : rgb ? [
        Number(rgb[1]),
        Number(rgb[2]),
        Number(rgb[3]),
        rgb[4] === void 0 ? 1 : Number(rgb[4])
      ] : void 0;
      if (!channels) return color;
      return `rgba(${channels.slice(0, 3).map((channel) => brightness ? Math.round(channel * (1 - Math.abs(brightness)) + Math.max(0, brightness) * 255) : channel).join(", ")}, ${Math.max(0, Math.min(channels[3] * opacity, 1))})`;
    }
    function animatedRunColor(color, opacity) {
      if (opacity >= 1) return color;
      const value = color || "currentColor";
      const adjusted = applyColorOpacity(value, opacity);
      return adjusted !== value || value === "transparent" ? adjusted : `color-mix(in srgb, ${value} ${Math.max(0, Math.min(opacity, 1)) * 100}%, transparent)`;
    }
    function elementStyle(element) {
      var _a, _b, _c;
      const fittedSize = autoFitDimensions.get(`${props.slide.id}:${element.id}`);
      const width = (fittedSize == null ? void 0 : fittedSize.width) || element.width;
      const height = (fittedSize == null ? void 0 : fittedSize.height) || element.height;
      const flips = [element.flipH ? "scaleX(-1)" : "", element.flipV ? "scaleY(-1)" : ""].filter(Boolean).join(" ");
      const matrix = element.renderMatrix;
      const [scaleX, scaleY] = animationScale(element);
      const [motionX, motionY] = animationMotion(element);
      const [slideX, slideY] = animationSlide(element);
      const hasScaleAnimation = scaleX !== 1 || scaleY !== 1;
      const rotation = animationRotation(element);
      const rotationDelta = rotation - element.rotation;
      const hasRotationAnimation = Math.abs(rotationDelta) > 1e-4;
      const translations = [motionX || motionY ? `translate(${motionX * scale.value}px,${motionY * scale.value}px)` : "", slideX || slideY ? `translate(${slideX * scale.value}px,${slideY * scale.value}px)` : ""].filter(Boolean).join(" ");
      const shadows = ((_a = element.shadows) == null ? void 0 : _a.length) ? element.shadows : element.shadow ? [element.shadow] : [];
      const fillOn = animationOnOff(element, "fill.on");
      const strokeOn = animationOnOff(element, "stroke.on");
      const shadowOn = animationOnOff(element, "shadow.on");
      const shadowColors = shadows.map((shadow) => shadowCss(shadow, applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity"))));
      const animatedFill = animationColor(element, "fillcolor");
      const fill = animatedFill || element.fill;
      const fillOpacity = animationOpacity(element, "fill.opacity");
      const stroke = applyColorOpacity(animationColor(element, "stroke.color") || element.stroke, animationOpacity(element, "stroke.opacity"));
      const strokeWidth = animationStrokeWidth(element);
      const strokeDash = animationStrokeDash(element);
      const glows = ((_b = element.glows) == null ? void 0 : _b.length) ? element.glows : element.glow ? [element.glow] : [];
      const customGeometry = Boolean((_c = element.customPaths) == null ? void 0 : _c.length);
      const radialFill = fillOn && !animatedFill && !customGeometry && !usesSvgPolygon(element) ? customSvgGradient(element, "fill") : void 0;
      const radialBackground = (radialFill == null ? void 0 : radialFill.kind) === "radial" && radialFill.fx !== void 0 && radialFill.fy !== void 0 && radialFill.fr !== void 0 ? radialGradientBackground(radialFill) : void 0;
      const showStroke = strokeOn && element.kind === "shape" && !customGeometry && !usesSvgOutline(element) && strokeWidth > 0 && stroke !== "transparent" && !stroke.includes("gradient");
      const transform = matrix ? [translations, `matrix(${matrix.slice(0, 4).join(",")},${matrix[4] * scale.value},${matrix[5] * scale.value})${hasScaleAnimation || hasRotationAnimation ? ` translate(${element.width * scale.value / 2}px,${element.height * scale.value / 2}px)${hasRotationAnimation ? ` rotate(${rotationDelta}deg)` : ""}${hasScaleAnimation ? ` scale(${scaleX},${scaleY})` : ""} translate(${-element.width * scale.value / 2}px,${-element.height * scale.value / 2}px)` : ""}`].filter(Boolean).join(" ") : [
        translations,
        rotation ? `rotate(${rotation}deg)` : "",
        flips,
        hasScaleAnimation ? `scale(${scaleX},${scaleY})` : ""
      ].filter(Boolean).join(" ");
      return {
        left: `${element.x / props.width * 100}%`,
        top: `${element.y / props.height * 100}%`,
        width: `${width / props.width * 100}%`,
        height: `${height / props.height * 100}%`,
        overflow: customGeometry || usesSvgOutline(element) ? "visible" : void 0,
        transform: transform || void 0,
        transformOrigin: matrix ? "0 0" : void 0,
        background: element.kind === "shape" && !customGeometry && !usesSvgPolygon(element) ? fillOn ? radialBackground || applyColorOpacity(fill, fillOpacity) : "none" : void 0,
        borderColor: showStroke ? stroke : void 0,
        borderWidth: showStroke ? `${strokeWidth * scale.value}px` : void 0,
        borderStyle: showStroke ? cssLineStyle(strokeDash) : void 0,
        boxShadow: shadows.length && shadowOn && !customGeometry && !usesSvgPolygon(element) ? shadowColors.join(", ") : void 0,
        filter: glows.length ? glows.map((glow) => `drop-shadow(0 0 ${glow.radius * scale.value}px ${glow.color})`).join(" ") : void 0
      };
    }
    function usesSvgOutline(element) {
      return element.kind === "shape" && ([
        "rect",
        "ellipse",
        "roundRect",
        "flowChartConnector",
        "flowChartSummingJunction",
        "flowChartOr",
        "flowChartProcess",
        "flowChartAlternateProcess",
        "flowChartDelay",
        "flowChartDisplay",
        "flowChartInternalStorage"
      ].includes(element.geometry) || usesSvgPolygon(element));
    }
    function usesSvgPolygon(element) {
      if (element.kind === "shape" && [
        "leftBracket",
        "rightBracket",
        "bracketPair",
        "leftBrace",
        "rightBrace",
        "bracePair"
      ].includes(element.geometry)) return true;
      if (element.kind === "shape" && [
        "cube",
        "plaque",
        "frame",
        "halfFrame",
        "cornerTabs",
        "plaqueTabs",
        "squareTabs",
        "donut",
        "pie",
        "chord",
        "arc",
        "blockArc",
        "pieWedge"
      ].includes(element.geometry)) return true;
      return element.kind === "shape" && [
        "plus",
        "mathMinus",
        "mathEqual",
        "mathPlus",
        "mathDivide",
        "mathMultiply",
        "mathNotEqual",
        "round1Rect",
        "round2DiagRect",
        "round2SameRect",
        "snipRoundRect",
        "snip1Rect",
        "snip2DiagRect",
        "snip2SameRect",
        "homePlate",
        "nonIsoscelesTrapezoid",
        "heptagon",
        "decagon",
        "leftRightArrow",
        "upDownArrow",
        "upArrow",
        "downArrow",
        "bentUpArrow",
        "quadArrow",
        "leftRightUpArrow",
        "rightArrow",
        "leftArrow",
        "notchedRightArrow",
        "heart",
        "lightningBolt",
        "triangle",
        "rtTriangle",
        "diamond",
        "flowChartDecision",
        "flowChartSort",
        "flowChartCollate",
        "flowChartPunchedCard",
        "flowChartOfflineStorage",
        "flowChartOnlineStorage",
        "flowChartPunchedTape",
        "flowChartMagneticDisk",
        "flowChartMagneticDrum",
        "flowChartMagneticTape",
        "flowChartMultidocument",
        "flowChartOffpageConnector",
        "flowChartDocument",
        "flowChartManualInput",
        "flowChartManualOperation",
        "flowChartInputOutput",
        "flowChartPreparation",
        "flowChartTerminator",
        "flowChartPredefinedProcess",
        "flowChartExtract",
        "flowChartMerge",
        "parallelogram",
        "trapezoid",
        "hexagon",
        "octagon",
        "pentagon",
        "star4",
        "star5",
        "star6",
        "star7",
        "star8",
        "star10",
        "star12",
        "star16",
        "star24",
        "star32",
        "dodecagon",
        "chevron"
      ].includes(element.geometry);
    }
    function presetBracketBracePaths(element) {
      var _a, _b, _c, _d, _e, _f, _g, _h;
      const rectangle = "M0 0 H1000 V1000 H0 Z";
      if (!(element.width > 0) || !(element.height > 0)) return [rectangle, rectangle];
      const shortSide = Math.min(element.width, element.height);
      if (element.geometry === "leftBrace" || element.geometry === "rightBrace") {
        const adjustment2 = Math.max(0, Math.min(1e5, (_b = (_a = element.presetAdjustments) == null ? void 0 : _a.adj2) != null ? _b : 5e4));
        const maximumAdjustment1 = Math.min(adjustment2, 1e5 - adjustment2) / 2 * element.height / shortSide;
        const radiusY = shortSide * Math.max(0, Math.min(maximumAdjustment1, (_d = (_c = element.presetAdjustments) == null ? void 0 : _c.adj1) != null ? _d : 8333)) / 1e5;
        const middleX = element.width / 2;
        const y3 = element.height * adjustment2 / 1e5;
        const y2 = y3 - radiusY, y4 = element.height - radiusY;
        const buildPath = (closed) => {
          let current = [0, 0];
          const point = ([x, y]) => `${Number((x / element.width * 1e3).toFixed(6))} ${Number((y / element.height * 1e3).toFixed(6))}`;
          const move = (x, y) => {
            current = [x, y];
            return `M${point(current)}`;
          };
          const line = (x, y) => {
            current = [x, y];
            return `L${point(current)}`;
          };
          const arc = (radiusX, startAngle, sweepAngle) => {
            if (!(radiusX > 0) || !(radiusY > 0)) return "";
            const start = startAngle * Math.PI / 180, end = (startAngle + sweepAngle) * Math.PI / 180;
            const centerX = current[0] - radiusX * Math.cos(start), centerY = current[1] - radiusY * Math.sin(start);
            current = [centerX + radiusX * Math.cos(end), centerY + radiusY * Math.sin(end)];
            return `A${Number((radiusX / element.width * 1e3).toFixed(6))} ${Number((radiusY / element.height * 1e3).toFixed(6))} 0 ${Math.abs(sweepAngle) > 180 ? 1 : 0} ${sweepAngle >= 0 ? 1 : 0} ${point(current)}`;
          };
          const commands = element.geometry === "leftBrace" ? [
            move(element.width, element.height),
            arc(middleX, 90, 90),
            line(middleX, y3 - radiusY),
            arc(middleX, 0, -90),
            arc(middleX, 90, -90),
            line(middleX, radiusY),
            arc(middleX, 180, 90)
          ] : [
            move(0, 0),
            arc(middleX, 270, 90),
            line(middleX, y2),
            arc(middleX, 180, -90),
            arc(middleX, 270, -90),
            line(middleX, y4),
            arc(middleX, 0, 90)
          ];
          if (closed) commands.push("Z");
          return commands.filter(Boolean).join(" ");
        };
        return [buildPath(true), buildPath(false)];
      }
      const isPair = element.geometry === "bracketPair";
      if (element.geometry === "bracePair") {
        const adjustment2 = Math.max(0, Math.min(25e3, (_f = (_e = element.presetAdjustments) == null ? void 0 : _e.adj) != null ? _f : 8333));
        const radius = shortSide * adjustment2 / 1e5;
        const x2 = shortSide * adjustment2 / 5e4;
        const x3 = element.width - x2, x4 = element.width - radius;
        const y2 = element.height / 2 - radius, y3 = element.height / 2 + radius, y4 = element.height - radius;
        const point = ([x, y]) => `${Number((x / element.width * 1e3).toFixed(6))} ${Number((y / element.height * 1e3).toFixed(6))}`;
        const buildPath = (outline2) => {
          let current = [0, 0];
          const move = (x, y) => {
            current = [x, y];
            return `M${point(current)}`;
          };
          const line = (x, y) => {
            current = [x, y];
            return `L${point(current)}`;
          };
          const arc = (startAngle, sweepAngle) => {
            if (!(radius > 0)) return "";
            const start = startAngle * Math.PI / 180, end = (startAngle + sweepAngle) * Math.PI / 180;
            const centerX = current[0] - radius * Math.cos(start), centerY = current[1] - radius * Math.sin(start);
            current = [centerX + radius * Math.cos(end), centerY + radius * Math.sin(end)];
            return `A${Number((radius / element.width * 1e3).toFixed(6))} ${Number((radius / element.height * 1e3).toFixed(6))} 0 ${Math.abs(sweepAngle) > 180 ? 1 : 0} ${sweepAngle >= 0 ? 1 : 0} ${point(current)}`;
          };
          const leftBrace = [
            move(x2, element.height),
            arc(90, 90),
            line(radius, y3),
            arc(0, -90),
            arc(90, -90),
            line(radius, radius),
            arc(180, 90)
          ];
          const rightBrace = [
            move(x3, 0),
            arc(270, 90),
            line(x4, y2),
            arc(180, -90),
            arc(270, -90),
            line(x4, y4),
            arc(0, 90)
          ];
          const joinAtTop = `L${point([x3, 0])}`;
          return (outline2 ? [...leftBrace, ...rightBrace] : [
            ...leftBrace,
            joinAtTop,
            ...rightBrace.slice(1),
            "Z"
          ]).filter(Boolean).join(" ");
        };
        return [buildPath(false), buildPath(true)];
      }
      const maximumAdjustment = isPair ? 5e4 : 5e4 * element.height / shortSide;
      const adjustment = Math.max(0, Math.min(maximumAdjustment, (_h = (_g = element.presetAdjustments) == null ? void 0 : _g.adj) != null ? _h : isPair ? 16667 : 8333));
      if (isPair) {
        const inset = shortSide * adjustment / 1e5 * 29289 / 1e5;
        const x = inset / element.width * 1e3, y = inset / element.height * 1e3;
        const oppositeX = 1e3 - x, oppositeY = 1e3 - y;
        return [`M0 ${y} A${x} ${y} 0 0 1 ${x} 0 L${oppositeX} 0 A${x} ${y} 0 0 1 1000 ${y} L1000 ${oppositeY} A${x} ${y} 0 0 1 ${oppositeX} 1000 L${x} 1000 A${x} ${y} 0 0 1 0 ${oppositeY} Z`, `M${x} 1000 A${x} ${y} 0 0 1 0 ${oppositeY} L0 ${y} A${x} ${y} 0 0 1 ${x} 0 M${oppositeX} 0 A${x} ${y} 0 0 1 1000 ${y} L1000 ${oppositeY} A${x} ${y} 0 0 1 ${oppositeX} 1000`];
      }
      const curveHeight = shortSide * adjustment / 1e5 / element.height * 1e3;
      const lowerCurveStart = 1e3 - curveHeight;
      const outline = element.geometry === "rightBracket" ? `M0 0 A1000 ${curveHeight} 0 0 1 1000 ${curveHeight} L1000 ${lowerCurveStart} A1000 ${curveHeight} 0 0 1 0 1000` : `M1000 1000 A1000 ${curveHeight} 0 0 1 0 ${lowerCurveStart} L0 ${curveHeight} A1000 ${curveHeight} 0 0 1 1000 0`;
      return [`${outline} Z`, outline];
    }
    function presetOutlineDetailPaths(element) {
      var _a;
      if (element.geometry === "leftBracket" || element.geometry === "rightBracket" || element.geometry === "bracketPair" || element.geometry === "leftBrace" || element.geometry === "rightBrace" || element.geometry === "bracePair") return [presetBracketBracePaths(element)[1]];
      if (element.geometry === "arc") {
        const arc = presetAngleArcPaths(element)[1];
        return arc ? [arc] : [];
      }
      if (element.geometry === "frame") return [presetOutlinePath(element)];
      if (element.geometry === "cube") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return [presetOutlinePath(element)];
        const inset = Math.min(width, height) * ((_a = element.presetAdjustment) != null ? _a : 0.25);
        const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
        return [
          presetOutlinePath(element),
          `M0 ${y(inset)} L${x(width - inset)} ${y(inset)} L1000 0`,
          `M${x(width - inset)} ${y(inset)} L${x(width - inset)} 1000`
        ];
      }
      if (element.geometry === "flowChartMagneticDisk") return [presetOutlinePath(element), "M1000 166.6667 A500 166.6667 0 0 1 0 166.6667"];
      if (element.geometry === "flowChartMagneticDrum") return [presetOutlinePath(element), `M833.3333333333334 1000 A${1e3 / 6} 500 0 0 1 833.3333333333334 0`];
      if (element.geometry === "flowChartMultidocument") {
        const scale2 = 1e3 / 21600;
        const point = (x, y) => `${x * scale2} ${y * scale2}`;
        return [
          `M${point(0, 3675)} L${point(18595, 3675)} L${point(18595, 18022)} C${point(9298, 18022)} ${point(9298, 23542)} ${point(0, 20782)} Z`,
          `M${point(1532, 3675)} L${point(1532, 1815)} L${point(2e4, 1815)} L${point(2e4, 16252)} C${point(19298, 16252)} ${point(18595, 16352)} ${point(18595, 16352)}`,
          `M${point(2972, 1815)} L${point(2972, 0)} L${point(21600, 0)} L${point(21600, 14392)} C${point(20800, 14392)} ${point(2e4, 14467)} ${point(2e4, 14467)}`
        ];
      }
      return [];
    }
    function roundRectRadii(element) {
      var _a;
      const fitted = autoFitDimensions.get(`${props.slide.id}:${element.id}`);
      const width = (fitted == null ? void 0 : fitted.width) || element.width, height = (fitted == null ? void 0 : fitted.height) || element.height;
      const radius = Math.min(width, height) * ((_a = element.presetAdjustment) != null ? _a : 0.16667);
      return [width > 0 ? radius / width * 1e3 : 0, height > 0 ? radius / height * 1e3 : 0];
    }
    function presetAngleArcPaths(element) {
      var _a, _b, _c, _d;
      const width = element.width, height = element.height;
      if (width <= 0 || height <= 0) return ["M0 0 H1000 V1000 H0 Z", ""];
      const fullTurn = 216e5;
      const startAngle = (_b = (_a = element.presetAdjustments) == null ? void 0 : _a.adj1) != null ? _b : element.geometry === "pie" ? 0 : element.geometry === "chord" ? 27e5 : 162e5;
      const endAngle = (_d = (_c = element.presetAdjustments) == null ? void 0 : _c.adj2) != null ? _d : element.geometry === "arc" ? 0 : 162e5;
      let sweepAngle = endAngle - startAngle;
      if (sweepAngle <= 0) sweepAngle += fullTurn;
      const [startX, startY] = presetEllipsePoint(width, height, startAngle);
      const arc = sweepAngle === fullTurn ? (() => {
        const [middleX, middleY] = presetEllipsePoint(width, height, startAngle + fullTurn / 2);
        return `M${startX} ${startY} A500 500 0 0 1 ${middleX} ${middleY} A500 500 0 0 1 ${startX} ${startY}`;
      })() : (() => {
        const [endX, endY] = presetEllipsePoint(width, height, endAngle);
        return `M${startX} ${startY} A500 500 0 ${sweepAngle > fullTurn / 2 ? 1 : 0} 1 ${endX} ${endY}`;
      })();
      return [element.geometry === "chord" ? `${arc} Z` : `${arc} L500 500 Z`, arc];
    }
    function presetBlockArcPath(element) {
      var _a, _b, _c, _d, _e, _f;
      const width = element.width, height = element.height;
      if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
      const fullTurn = 216e5;
      const startAngle = (_b = (_a = element.presetAdjustments) == null ? void 0 : _a.adj1) != null ? _b : 108e5;
      const endAngle = (_d = (_c = element.presetAdjustments) == null ? void 0 : _c.adj2) != null ? _d : 0;
      let sweepAngle = endAngle - startAngle;
      if (sweepAngle <= 0) sweepAngle += fullTurn;
      const inset = Math.min(width, height) * ((_f = (_e = element.presetAdjustments) == null ? void 0 : _e.adj3) != null ? _f : 25e3) / 1e5;
      const innerWidth = Math.max(0, width - 2 * inset);
      const innerHeight = Math.max(0, height - 2 * inset);
      const innerPoint = (angle) => {
        const [x, y] = presetEllipsePoint(innerWidth, innerHeight, angle);
        return [500 + (x - 500) * innerWidth / width, 500 + (y - 500) * innerHeight / height];
      };
      const arcCommands = (fromAngle, sweep, rx, ry, pointAt, direction) => {
        const [startX, startY] = pointAt(fromAngle);
        if (Math.abs(sweep) === fullTurn) {
          const [middleX, middleY] = pointAt(fromAngle + sweep / 2);
          return `A${rx} ${ry} 0 0 ${direction} ${middleX} ${middleY} A${rx} ${ry} 0 0 ${direction} ${startX} ${startY}`;
        }
        const [endX, endY] = pointAt(fromAngle + sweep);
        return `A${rx} ${ry} 0 ${Math.abs(sweep) > fullTurn / 2 ? 1 : 0} ${direction} ${endX} ${endY}`;
      };
      const [outerStartX, outerStartY] = presetEllipsePoint(width, height, startAngle);
      const [innerEndX, innerEndY] = innerPoint(endAngle);
      return `M${outerStartX} ${outerStartY} ${arcCommands(startAngle, sweepAngle, 500, 500, (angle) => presetEllipsePoint(width, height, angle), 1)} L${innerEndX} ${innerEndY} ${arcCommands(endAngle, -sweepAngle, 500 * innerWidth / width, 500 * innerHeight / height, innerPoint, 0)} Z`;
    }
    function presetOutlinePath(element) {
      var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y, _z, _A, _B, _C, _D, _E, _F, _G, _H, _I, _J, _K, _L, _M, _N, _O, _P, _Q, _R, _S, _T, _U, _V, _W, _X, _Y, _Z, __, _$, _aa, _ba, _ca, _da, _ea, _fa, _ga, _ha, _ia, _ja, _ka, _la, _ma, _na, _oa, _pa, _qa, _ra, _sa, _ta, _ua, _va, _wa, _xa, _ya, _za, _Aa, _Ba, _Ca, _Da, _Ea, _Fa, _Ga, _Ha, _Ia, _Ja, _Ka, _La, _Ma, _Na, _Oa, _Pa, _Qa, _Ra, _Sa, _Ta, _Ua, _Va, _Wa, _Xa, _Ya, _Za, __a, _$a, _ab, _bb, _cb, _db, _eb, _fb, _gb, _hb, _ib, _jb, _kb, _lb, _mb, _nb, _ob, _pb, _qb, _rb, _sb, _tb, _ub, _vb, _wb, _xb, _yb, _zb, _Ab, _Bb, _Cb;
      if (element.geometry === "leftBracket" || element.geometry === "rightBracket" || element.geometry === "bracketPair" || element.geometry === "leftBrace" || element.geometry === "rightBrace" || element.geometry === "bracePair") return presetBracketBracePaths(element)[0];
      if (element.geometry === "quadArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj2 = Math.max(0, Math.min(5e4, (_b = (_a = element.presetAdjustments) == null ? void 0 : _a.adj2) != null ? _b : 22500));
        const adj1 = Math.max(0, Math.min(2 * adj2, (_d = (_c = element.presetAdjustments) == null ? void 0 : _c.adj1) != null ? _d : 22500));
        const x1 = shortSide * Math.max(0, Math.min((1e5 - 2 * adj2) / 2, (_f = (_e = element.presetAdjustments) == null ? void 0 : _e.adj3) != null ? _f : 22500)) / 1e5;
        const dx2 = shortSide * adj2 / 1e5;
        const dx3 = shortSide * adj1 / 2e5;
        const x2 = width / 2 - dx2, x3 = width / 2 - dx3, x4 = width / 2 + dx3, x5 = width / 2 + dx2, x6 = width - x1;
        const y2 = height / 2 - dx2, y3 = height / 2 - dx3, y4 = height / 2 + dx3, y5 = height / 2 + dx2, y6 = height - x1;
        const x7 = (value) => value / width * 1e3, y7 = (value) => value / height * 1e3;
        return `M0 500 L${x7(x1)} ${y7(y2)} L${x7(x1)} ${y7(y3)} L${x7(x3)} ${y7(y3)} L${x7(x3)} ${y7(x1)} L${x7(x2)} ${y7(x1)} L500 0 L${x7(x5)} ${y7(x1)} L${x7(x4)} ${y7(x1)} L${x7(x4)} ${y7(y3)} L${x7(x6)} ${y7(y3)} L${x7(x6)} ${y7(y2)} L1000 500 L${x7(x6)} ${y7(y5)} L${x7(x6)} ${y7(y4)} L${x7(x4)} ${y7(y4)} L${x7(x4)} ${y7(y6)} L${x7(x5)} ${y7(y6)} L500 1000 L${x7(x2)} ${y7(y6)} L${x7(x3)} ${y7(y6)} L${x7(x3)} ${y7(y4)} L${x7(x1)} ${y7(y4)} L${x7(x1)} ${y7(y5)} Z`;
      }
      if (element.geometry === "leftRightUpArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj2 = Math.max(0, Math.min(5e4, (_h = (_g = element.presetAdjustments) == null ? void 0 : _g.adj2) != null ? _h : 25e3));
        const adj1 = Math.max(0, Math.min(2 * adj2, (_j = (_i = element.presetAdjustments) == null ? void 0 : _i.adj1) != null ? _j : 25e3));
        const x1 = shortSide * Math.max(0, Math.min((1e5 - 2 * adj2) / 2, (_l = (_k = element.presetAdjustments) == null ? void 0 : _k.adj3) != null ? _l : 25e3)) / 1e5;
        const dx2 = shortSide * adj2 / 1e5;
        const dx3 = shortSide * adj1 / 2e5;
        const x2 = width / 2 - dx2, x3 = width / 2 - dx3, x4 = width / 2 + dx3, x5 = width / 2 + dx2, x6 = width - x1;
        const y2 = height - shortSide * adj2 / 5e4, y4 = height - dx2, y3 = y4 - dx3, y5 = y4 + dx3;
        const x7 = (value) => value / width * 1e3, y6 = (value) => value / height * 1e3;
        return `M0 ${y6(y4)} L${x7(x1)} ${y6(y2)} L${x7(x1)} ${y6(y3)} L${x7(x3)} ${y6(y3)} L${x7(x3)} ${y6(x1)} L${x7(x2)} ${y6(x1)} L500 0 L${x7(x5)} ${y6(x1)} L${x7(x4)} ${y6(x1)} L${x7(x4)} ${y6(y3)} L${x7(x6)} ${y6(y3)} L${x7(x6)} ${y6(y2)} L1000 ${y6(y4)} L${x7(x6)} 1000 L${x7(x6)} ${y6(y5)} L${x7(x1)} ${y6(y5)} L${x7(x1)} 1000 Z`;
      }
      if (element.geometry === "bentUpArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(5e4, (_n = (_m = element.presetAdjustments) == null ? void 0 : _m.adj1) != null ? _n : 25e3));
        const adj2 = Math.max(0, Math.min(5e4, (_p = (_o = element.presetAdjustments) == null ? void 0 : _o.adj2) != null ? _p : 25e3));
        const adj3 = Math.max(0, Math.min(5e4, (_r = (_q = element.presetAdjustments) == null ? void 0 : _q.adj3) != null ? _r : 25e3));
        const dx2 = shortSide * adj1 / 2e5;
        const x3 = width - shortSide * adj1 / 1e5;
        const x2 = x3 - dx2, x4 = x3 + dx2;
        const y1 = shortSide * adj3 / 1e5;
        const y2 = height - shortSide * adj1 / 1e5;
        const x1 = width - shortSide * adj2 / 5e4;
        const x5 = (value) => value / width * 1e3, y3 = (value) => value / height * 1e3;
        return `M0 ${y3(y2)} L${x5(x2)} ${y3(y2)} L${x5(x2)} ${y3(y1)} L${x5(x1)} ${y3(y1)} L${x5(x3)} 0 L1000 ${y3(y1)} L${x5(x4)} ${y3(y1)} V1000 H0 Z`;
      }
      if (element.geometry === "upArrow" || element.geometry === "downArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(1e5, (_t = (_s = element.presetAdjustments) == null ? void 0 : _s.adj1) != null ? _t : 5e4));
        const adj2 = Math.max(0, Math.min(1e5 * height / shortSide, (_v = (_u = element.presetAdjustments) == null ? void 0 : _u.adj2) != null ? _v : 5e4));
        const halfHeadWidth = width * adj1 / 2e5;
        const x1 = width / 2 - halfHeadWidth, x2 = width / 2 + halfHeadWidth;
        const arrowDepth = shortSide * adj2 / 1e5;
        const x3 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        if (element.geometry === "upArrow") {
          const y22 = arrowDepth;
          return `M0 ${y2(y22)} L500 0 L1000 ${y2(y22)} L${x3(x2)} ${y2(y22)} V1000 H${x3(x1)} V${y2(y22)} Z`;
        }
        const y1 = height - arrowDepth;
        return `M0 ${y2(y1)} L${x3(x1)} ${y2(y1)} V0 H${x3(x2)} V${y2(y1)} H1000 L500 1000 Z`;
      }
      if (element.geometry === "leftArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(1e5, (_x = (_w = element.presetAdjustments) == null ? void 0 : _w.adj1) != null ? _x : 5e4));
        const x2 = shortSide * Math.max(0, Math.min(1e5 * width / shortSide, (_z = (_y = element.presetAdjustments) == null ? void 0 : _y.adj2) != null ? _z : 5e4)) / 1e5;
        const arrowHalfHeight = height * adj1 / 2e5;
        const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
        const x3 = (value) => value / width * 1e3, y3 = (value) => value / height * 1e3;
        return `M0 500 L${x3(x2)} 0 L${x3(x2)} ${y3(y1)} H1000 V${y3(y2)} H${x3(x2)} L${x3(x2)} 1000 Z`;
      }
      if (element.geometry === "notchedRightArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(1e5, (_B = (_A = element.presetAdjustments) == null ? void 0 : _A.adj1) != null ? _B : 5e4));
        const arrowLength = shortSide * Math.max(0, Math.min(1e5 * width / shortSide, (_D = (_C = element.presetAdjustments) == null ? void 0 : _C.adj2) != null ? _D : 5e4)) / 1e5;
        const x2 = width - arrowLength;
        const arrowHalfHeight = height * adj1 / 2e5;
        const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
        const x1 = arrowHalfHeight * arrowLength / (height / 2);
        const x3 = (value) => value / width * 1e3, y3 = (value) => value / height * 1e3;
        return `M0 ${y3(y1)} L${x3(x2)} ${y3(y1)} L${x3(x2)} 0 L1000 500 L${x3(x2)} 1000 L${x3(x2)} ${y3(y2)} H0 L${x3(x1)} 500 Z`;
      }
      if (element.geometry === "rightArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(1e5, (_F = (_E = element.presetAdjustments) == null ? void 0 : _E.adj1) != null ? _F : 5e4));
        const x1 = width - shortSide * Math.max(0, Math.min(1e5 * width / shortSide, (_H = (_G = element.presetAdjustments) == null ? void 0 : _G.adj2) != null ? _H : 5e4)) / 1e5;
        const arrowHalfHeight = height * adj1 / 2e5;
        const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
        const x2 = (value) => value / width * 1e3, y3 = (value) => value / height * 1e3;
        return `M0 ${y3(y1)} L${x2(x1)} ${y3(y1)} L${x2(x1)} 0 L1000 500 L${x2(x1)} 1000 L${x2(x1)} ${y3(y2)} H0 Z`;
      }
      if (element.geometry === "upDownArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(1e5, (_J = (_I = element.presetAdjustments) == null ? void 0 : _I.adj1) != null ? _J : 5e4));
        const adj2 = Math.max(0, Math.min(5e4 * height / shortSide, (_L = (_K = element.presetAdjustments) == null ? void 0 : _K.adj2) != null ? _L : 5e4));
        const arrowHalfWidth = width * adj1 / 2e5;
        const x1 = width / 2 - arrowHalfWidth, x2 = width / 2 + arrowHalfWidth;
        const arrowHeight = shortSide * adj2 / 1e5;
        const y2 = arrowHeight, y3 = height - arrowHeight;
        const x3 = (value) => value / width * 1e3, y4 = (value) => value / height * 1e3;
        return `M500 0 L1000 ${y4(y2)} L${x3(x2)} ${y4(y2)} L${x3(x2)} ${y4(y3)} L1000 ${y4(y3)} L500 1000 L0 ${y4(y3)} L${x3(x1)} ${y4(y3)} L${x3(x1)} ${y4(y2)} L0 ${y4(y2)} Z`;
      }
      if (element.geometry === "leftRightArrow") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const adj1 = Math.max(0, Math.min(1e5, (_N = (_M = element.presetAdjustments) == null ? void 0 : _M.adj1) != null ? _N : 5e4));
        const arrowWidth = shortSide * Math.max(0, Math.min(5e4 * width / shortSide, (_P = (_O = element.presetAdjustments) == null ? void 0 : _O.adj2) != null ? _P : 5e4)) / 1e5;
        const x2 = arrowWidth, x3 = width - arrowWidth;
        const arrowHalfHeight = height * adj1 / 2e5;
        const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
        const x4 = (value) => value / width * 1e3, y3 = (value) => value / height * 1e3;
        return `M0 500 L${x4(x2)} 0 L${x4(x2)} ${y3(y1)} L${x4(x3)} ${y3(y1)} L${x4(x3)} 0 L1000 500 L${x4(x3)} 1000 L${x4(x3)} ${y3(y2)} L${x4(x2)} ${y3(y2)} L${x4(x2)} 1000 Z`;
      }
      if (element.geometry === "plus") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const inset = Math.min(width, height) * Math.max(0, Math.min(0.5, (_Q = element.presetAdjustment) != null ? _Q : 0.25));
        const x1 = inset / width * 1e3, x2 = 1e3 - x1;
        const y1 = inset / height * 1e3, y2 = 1e3 - y1;
        return `M0 ${y1} L${x1} ${y1} L${x1} 0 L${x2} 0 L${x2} ${y1} H1000 V${y2} H${x2} V1000 H${x1} V${y2} H0 Z`;
      }
      if (element.geometry === "mathMinus") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const adj1 = Math.max(0, Math.min(1e5, (_S = (_R = element.presetAdjustments) == null ? void 0 : _R.adj1) != null ? _S : 23520));
        const halfBarWidth = width * 73490 / 2e5;
        const halfBarHeight = height * adj1 / 2e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(width / 2 - halfBarWidth)} ${y2(height / 2 - halfBarHeight)} H${x2(width / 2 + halfBarWidth)} V${y2(height / 2 + halfBarHeight)} H${x2(width / 2 - halfBarWidth)} Z`;
      }
      if (element.geometry === "mathEqual") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const adj1 = Math.max(0, Math.min(36745, (_U = (_T = element.presetAdjustments) == null ? void 0 : _T.adj1) != null ? _U : 23520));
        const adj2 = Math.max(0, Math.min(1e5 - 2 * adj1, (_W = (_V = element.presetAdjustments) == null ? void 0 : _V.adj2) != null ? _W : 11760));
        const halfBarWidth = width * 73490 / 2e5;
        const halfGap = height * adj2 / 2e5;
        const barHeight = height * adj1 / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        const left = x2(width / 2 - halfBarWidth), right = x2(width / 2 + halfBarWidth);
        const upperInner = y2(height / 2 - halfGap);
        return `M${left} ${y2(height / 2 - halfGap - barHeight)} H${right} V${upperInner} H${left} Z M${left} ${y2(height / 2 + halfGap)} H${right} V${y2(height / 2 + halfGap + barHeight)} H${left} Z`;
      }
      if (element.geometry === "mathPlus") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const adj1 = Math.max(0, Math.min(73490, (_Y = (_X = element.presetAdjustments) == null ? void 0 : _X.adj1) != null ? _Y : 23520));
        const longHalfArm = width * 73490 / 2e5;
        const crossHalfArm = Math.min(width, height) * adj1 / 2e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        const x1 = x2(width / 2 - longHalfArm), x22 = x2(width / 2 - crossHalfArm), x3 = x2(width / 2 + crossHalfArm), x4 = x2(width / 2 + longHalfArm);
        const y1 = y2(height / 2 - height * 73490 / 2e5), y22 = y2(height / 2 - crossHalfArm), y3 = y2(height / 2 + crossHalfArm), y4 = y2(height / 2 + height * 73490 / 2e5);
        return `M${x1} ${y22} L${x22} ${y22} L${x22} ${y1} L${x3} ${y1} L${x3} ${y22} L${x4} ${y22} L${x4} ${y3} L${x3} ${y3} L${x3} ${y4} L${x22} ${y4} L${x22} ${y3} L${x1} ${y3} Z`;
      }
      if (element.geometry === "mathDivide") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const adj1 = Math.max(1e3, Math.min(36745, (__ = (_Z = element.presetAdjustments) == null ? void 0 : _Z.adj1) != null ? __ : 23520));
        const maximumAdj3 = Math.min((73490 + adj1) / 4, 36745 * width / height);
        const adj3 = Math.max(1e3, Math.min(maximumAdj3, (_aa = (_$ = element.presetAdjustments) == null ? void 0 : _$.adj3) != null ? _aa : 11760));
        const maximumAdj2 = 73490 - 4 * adj3 - adj1;
        const adj2 = Math.max(0, Math.min(maximumAdj2, (_ca = (_ba = element.presetAdjustments) == null ? void 0 : _ba.adj2) != null ? _ca : 5880));
        const dy1 = height * adj1 / 2e5, gap = height * adj2 / 1e5, radius = height * adj3 / 1e5;
        const y3 = height / 2 - dy1, y4 = height / 2 + dy1;
        const y1 = y3 - gap, y2 = y1 + radius, y5 = height - y1;
        const halfBarWidth = width * 73490 / 2e5;
        const x2 = (value) => value / width * 1e3, y6 = (value) => value / height * 1e3;
        const circle = (centerY) => `M${x2(width / 2)} ${y6(centerY - radius)} A${x2(radius)} ${y6(radius)} 0 1 1 ${x2(width / 2)} ${y6(centerY + radius)} A${x2(radius)} ${y6(radius)} 0 1 1 ${x2(width / 2)} ${y6(centerY - radius)} Z`;
        return `${circle(y2)} ${circle(y5 - radius)} M${x2(width / 2 - halfBarWidth)} ${y6(y3)} H${x2(width / 2 + halfBarWidth)} V${y6(y4)} H${x2(width / 2 - halfBarWidth)} Z`;
      }
      if (element.geometry === "mathMultiply") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const adj1 = Math.max(0, Math.min(51965, (_ea = (_da = element.presetAdjustments) == null ? void 0 : _da.adj1) != null ? _ea : 23520));
        const angle = Math.atan2(height, width), tangent = Math.tan(angle);
        const widthAlongDiagonal = Math.hypot(width, height) * (1 - 0.51965);
        const thickness = Math.min(width, height) * adj1 / 1e5;
        const xM = Math.cos(angle) * widthAlongDiagonal / 2, yM = Math.sin(angle) * widthAlongDiagonal / 2;
        const dx = Math.sin(angle) * thickness / 2, dy = Math.cos(angle) * thickness / 2;
        const xA = xM - dx, yA = yM + dy, xB = xM + dx, yB = yM - dy;
        const yC = (width / 2 - xB) * tangent + yB;
        const xD = width - xB, xE = width - xA;
        const xFE = tangent ? (height / 2 - yA) / tangent : 0;
        const xF = xE - xFE, xL = xA + xFE;
        const yG = height - yA, yH = height - yB, yI = height - yC;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(xA)} ${y2(yA)} L${x2(xB)} ${y2(yB)} L500 ${y2(yC)} L${x2(xD)} ${y2(yB)} L${x2(xE)} ${y2(yA)} L${x2(xF)} 500 L${x2(xE)} ${y2(yG)} L${x2(xD)} ${y2(yH)} L500 ${y2(yI)} L${x2(xB)} ${y2(yH)} L${x2(xA)} ${y2(yG)} L${x2(xL)} 500 Z`;
      }
      if (element.geometry === "mathNotEqual") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const adj1 = Math.max(0, Math.min(5e4, (_ga = (_fa = element.presetAdjustments) == null ? void 0 : _fa.adj1) != null ? _ga : 23520));
        const angle = Math.max(42e5, Math.min(66e5, (_ia = (_ha = element.presetAdjustments) == null ? void 0 : _ha.adj2) != null ? _ia : 66e5)) - 54e5;
        const adj3 = Math.max(0, Math.min(1e5 - 2 * adj1, (_ka = (_ja = element.presetAdjustments) == null ? void 0 : _ja.adj3) != null ? _ka : 11760));
        const halfHeight = height / 2;
        const dy1 = height * adj1 / 1e5, dy2 = height * adj3 / 2e5;
        const x1 = width / 2 - width * 73490 / 2e5, x8 = width / 2 + width * 73490 / 2e5;
        const y2 = halfHeight - dy2, y3 = halfHeight + dy2, y1 = y2 - dy1, y4 = y3 + dy1;
        const xAdj = halfHeight * Math.tan(angle / 6e4 * Math.PI / 180);
        const length = Math.hypot(xAdj, halfHeight);
        const barHalfWidth = length * dy1 / halfHeight / 2;
        const x7 = width / 2 + xAdj - barHalfWidth;
        const x6 = x7 - xAdj * y1 / halfHeight, x5 = x7 - xAdj * y2 / halfHeight;
        const x4 = x7 - xAdj * y3 / halfHeight, x3 = x7 - xAdj * y4 / halfHeight;
        const rx7 = x7 - barHalfWidth, rx6 = x6 - barHalfWidth;
        const rx5 = x5 - barHalfWidth, rx4 = x4 - barHalfWidth, rx3 = x3 - barHalfWidth;
        const dx7 = dy1 * halfHeight / length;
        const dy3 = dy1 * xAdj / length;
        const rx = angle > 0 ? x7 + dx7 : rx7;
        const lx = angle > 0 ? x7 : rx7 - dx7;
        const ry = angle > 0 ? dy3 : 0;
        const ly = angle > 0 ? 0 : -dy3;
        const dlx = width - rx, drx = width - lx, dly = height - ry, dry = height - ly;
        const x2 = (value) => value / width * 1e3, y5 = (value) => value / height * 1e3;
        return `M${x2(x1)} ${y5(y1)} L${x2(x6)} ${y5(y1)} L${x2(lx)} ${y5(ly)} L${x2(rx)} ${y5(ry)} L${x2(rx6)} ${y5(y1)} L${x2(x8)} ${y5(y1)} L${x2(x8)} ${y5(y2)} L${x2(rx5)} ${y5(y2)} L${x2(rx4)} ${y5(y3)} L${x2(x8)} ${y5(y3)} L${x2(x8)} ${y5(y4)} L${x2(rx3)} ${y5(y4)} L${x2(drx)} ${y5(dry)} L${x2(dlx)} ${y5(dly)} L${x2(x3)} ${y5(y4)} L${x2(x1)} ${y5(y4)} L${x2(x1)} ${y5(y3)} L${x2(x4)} ${y5(y3)} L${x2(x5)} ${y5(y2)} L${x2(x1)} ${y5(y2)} Z`;
      }
      if (element.geometry === "snip1Rect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const inset = Math.min(width, height) * ((_la = element.presetAdjustment) != null ? _la : 0.16667);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 0 H${x2(width - inset)} L1000 ${y2(inset)} V1000 H0 Z`;
      }
      if (element.geometry === "snip2DiagRect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const leftCut = shortSide * Math.max(0, Math.min(5e4, (_na = (_ma = element.presetAdjustments) == null ? void 0 : _ma.adj1) != null ? _na : 0)) / 1e5;
        const rightCut = shortSide * Math.max(0, Math.min(5e4, (_pa = (_oa = element.presetAdjustments) == null ? void 0 : _oa.adj2) != null ? _pa : 16667)) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(leftCut)} 0 L${x2(width - rightCut)} 0 L1000 ${y2(rightCut)} L1000 ${y2(height - leftCut)} L${x2(width - leftCut)} 1000 L${x2(rightCut)} 1000 L0 ${y2(height - rightCut)} L0 ${y2(leftCut)} Z`;
      }
      if (element.geometry === "snip2SameRect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const topCut = shortSide * Math.max(0, Math.min(5e4, (_ra = (_qa = element.presetAdjustments) == null ? void 0 : _qa.adj1) != null ? _ra : 16667)) / 1e5;
        const bottomCut = shortSide * Math.max(0, Math.min(5e4, (_ta = (_sa = element.presetAdjustments) == null ? void 0 : _sa.adj2) != null ? _ta : 0)) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(topCut)} 0 L${x2(width - topCut)} 0 L1000 ${y2(topCut)} L1000 ${y2(height - bottomCut)} L${x2(width - bottomCut)} 1000 L${x2(bottomCut)} 1000 L0 ${y2(height - bottomCut)} L0 ${y2(topCut)} Z`;
      }
      if (element.geometry === "round1Rect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const radius = Math.min(width, height) * ((_ua = element.presetAdjustment) != null ? _ua : 0.16667);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 0 L${x2(width - radius)} 0 A${x2(radius)} ${y2(radius)} 0 0 1 1000 ${y2(radius)} L1000 1000 H0 Z`;
      }
      if (element.geometry === "round2DiagRect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const radius1 = shortSide * ((_wa = (_va = element.presetAdjustments) == null ? void 0 : _va.adj1) != null ? _wa : 16667) / 1e5;
        const radius2 = shortSide * ((_ya = (_xa = element.presetAdjustments) == null ? void 0 : _xa.adj2) != null ? _ya : 0) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(radius1)} 0 L${x2(width - radius2)} 0 A${x2(radius2)} ${y2(radius2)} 0 0 1 1000 ${y2(radius2)} L1000 ${y2(height - radius1)} A${x2(radius1)} ${y2(radius1)} 0 0 1 ${x2(width - radius1)} 1000 L${x2(radius2)} 1000 A${x2(radius2)} ${y2(radius2)} 0 0 1 0 ${y2(height - radius2)} L0 ${y2(radius1)} A${x2(radius1)} ${y2(radius1)} 0 0 1 ${x2(radius1)} 0 Z`;
      }
      if (element.geometry === "round2SameRect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const topRadius = shortSide * ((_Aa = (_za = element.presetAdjustments) == null ? void 0 : _za.adj1) != null ? _Aa : 16667) / 1e5;
        const bottomRadius = shortSide * ((_Ca = (_Ba = element.presetAdjustments) == null ? void 0 : _Ba.adj2) != null ? _Ca : 0) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(topRadius)} 0 L${x2(width - topRadius)} 0 A${x2(topRadius)} ${y2(topRadius)} 0 0 1 1000 ${y2(topRadius)} L1000 ${y2(height - bottomRadius)} A${x2(bottomRadius)} ${y2(bottomRadius)} 0 0 1 ${x2(width - bottomRadius)} 1000 L${x2(bottomRadius)} 1000 A${x2(bottomRadius)} ${y2(bottomRadius)} 0 0 1 0 ${y2(height - bottomRadius)} L0 ${y2(topRadius)} A${x2(topRadius)} ${y2(topRadius)} 0 0 1 ${x2(topRadius)} 0 Z`;
      }
      if (element.geometry === "snipRoundRect") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const radius = shortSide * ((_Ea = (_Da = element.presetAdjustments) == null ? void 0 : _Da.adj1) != null ? _Ea : 16667) / 1e5;
        const cut = shortSide * ((_Ga = (_Fa = element.presetAdjustments) == null ? void 0 : _Fa.adj2) != null ? _Ga : 16667) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(radius)} 0 L${x2(width - cut)} 0 L1000 ${y2(cut)} L1000 1000 H0 L0 ${y2(radius)} A${x2(radius)} ${y2(radius)} 0 0 1 ${x2(radius)} 0 Z`;
      }
      if (element.geometry === "cube") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const inset = Math.min(width, height) * ((_Ha = element.presetAdjustment) != null ? _Ha : 0.25);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 ${y2(inset)} L${x2(inset)} 0 H1000 V${y2(height - inset)} L${x2(width - inset)} 1000 H0 Z`;
      }
      if (element.geometry === "frame") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const inset = Math.min(width, height) * Math.max(0, Math.min(5e4, (_Ja = (_Ia = element.presetAdjustments) == null ? void 0 : _Ia.adj1) != null ? _Ja : 12500)) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 0 H1000 V1000 H0 Z M${x2(inset)} ${y2(inset)} V${y2(height - inset)} H${x2(width - inset)} V${y2(inset)} Z`;
      }
      if (element.geometry === "halfFrame") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const maximumAdj2 = 1e5 * width / shortSide;
        const x1 = shortSide * Math.max(0, Math.min(maximumAdj2, (_La = (_Ka = element.presetAdjustments) == null ? void 0 : _Ka.adj2) != null ? _La : 33333)) / 1e5;
        const maximumAdj1 = 1e5 * (height - height * x1 / width) / shortSide;
        const y1 = shortSide * Math.max(0, Math.min(maximumAdj1, (_Na = (_Ma = element.presetAdjustments) == null ? void 0 : _Ma.adj1) != null ? _Na : 33333)) / 1e5;
        const x2 = width - y1 * width / height;
        const y2 = height - x1 * height / width;
        const x3 = (value) => value / width * 1e3, y3 = (value) => value / height * 1e3;
        return `M0 0 H1000 L${x3(x2)} ${y3(y1)} L${x3(x1)} ${y3(y1)} L${x3(x1)} ${y3(y2)} Z`;
      }
      if (element.geometry === "cornerTabs" || element.geometry === "plaqueTabs" || element.geometry === "squareTabs") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "";
        const inset = Math.hypot(width, height) / 20;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        const rightInset = x2(width - inset), bottomInset = y2(height - inset), topRight = y2(inset), leftRight = x2(inset);
        if (element.geometry === "cornerTabs") return `M0 0 L${leftRight} 0 L0 ${topRight} Z M0 ${bottomInset} L${leftRight} 1000 L0 1000 Z M${rightInset} 0 L1000 0 L1000 ${topRight} Z M1000 ${bottomInset} L1000 1000 L${rightInset} 1000 Z`;
        if (element.geometry === "squareTabs") return `M0 0 L${leftRight} 0 L${leftRight} ${topRight} L0 ${topRight} Z M0 ${bottomInset} L${leftRight} ${bottomInset} L${leftRight} 1000 L0 1000 Z M${rightInset} 0 L1000 0 L1000 ${topRight} L${rightInset} ${topRight} Z M${rightInset} ${bottomInset} L1000 ${bottomInset} L1000 1000 L${rightInset} 1000 Z`;
        return `M0 0 L${leftRight} 0 A${leftRight} ${topRight} 0 0 1 0 ${topRight} Z M0 ${bottomInset} A${leftRight} ${topRight} 0 0 1 ${leftRight} 1000 L0 1000 Z M1000 0 L1000 ${topRight} A${leftRight} ${topRight} 0 0 1 ${rightInset} 0 Z M${rightInset} 1000 A${leftRight} ${topRight} 0 0 1 1000 ${bottomInset} L1000 1000 Z`;
      }
      if (element.geometry === "plaque") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const radius = Math.min(width, height) * ((_Oa = element.presetAdjustment) != null ? _Oa : 0.16667);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 ${y2(radius)} A${x2(radius)} ${y2(radius)} 0 0 1 ${x2(radius)} 0 L${x2(width - radius)} 0 A${x2(radius)} ${y2(radius)} 0 0 1 1000 ${y2(radius)} L1000 ${y2(height - radius)} A${x2(radius)} ${y2(radius)} 0 0 1 ${x2(width - radius)} 1000 L${x2(radius)} 1000 A${x2(radius)} ${y2(radius)} 0 0 1 0 ${y2(height - radius)} Z`;
      }
      if (element.geometry === "donut") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const inset = Math.min(width, height) * ((_Pa = element.presetAdjustment) != null ? _Pa : 0.25);
        const innerRx = (width / 2 - inset) / width * 1e3, innerRy = (height / 2 - inset) / height * 1e3;
        return `M0 500 A500 500 0 0 1 500 0 A500 500 0 0 1 1000 500 A500 500 0 0 1 500 1000 A500 500 0 0 1 0 500 Z ${innerRx > 0 && innerRy > 0 ? `M${inset / width * 1e3} 500 A${innerRx} ${innerRy} 0 0 0 500 ${500 + innerRy} A${innerRx} ${innerRy} 0 0 0 ${1e3 - inset / width * 1e3} 500 A${innerRx} ${innerRy} 0 0 0 500 ${500 - innerRy} A${innerRx} ${innerRy} 0 0 0 ${inset / width * 1e3} 500 Z` : ""}`;
      }
      if ([
        "pie",
        "chord",
        "arc"
      ].includes(element.geometry)) return presetAngleArcPaths(element)[0];
      if (element.geometry === "blockArc") return presetBlockArcPath(element);
      if (element.geometry === "pieWedge") return "M0 1000 A1000 1000 0 0 1 1000 0 L1000 1000 Z";
      if (element.geometry === "homePlate") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const x2 = (width - Math.min(width, height) * ((_Qa = element.presetAdjustment) != null ? _Qa : 0.5)) / width * 1e3;
        return `M0 0 L${x2} 0 L1000 500 L${x2} 1000 L0 1000 Z`;
      }
      if (element.geometry === "nonIsoscelesTrapezoid") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const shortSide = Math.min(width, height);
        const leftTop = shortSide * ((_Sa = (_Ra = element.presetAdjustments) == null ? void 0 : _Ra.adj1) != null ? _Sa : 25e3) / 1e5;
        const rightTop = width - shortSide * ((_Ua = (_Ta = element.presetAdjustments) == null ? void 0 : _Ta.adj2) != null ? _Ua : 25e3) / 1e5;
        const x2 = (value) => value / width * 1e3;
        return `M0 1000 L${x2(leftTop)} 0 L${x2(rightTop)} 0 L1000 1000 Z`;
      }
      if (element.geometry === "heart") {
        const innerControlX = 500 + 1e4 / 48, outerControlX = 500 + 49e3 / 48;
        const upperControlY = -1e3 / 3;
        return `M500 250 C${innerControlX} ${upperControlY} ${outerControlX} 250 500 1000 C-520.8333333333335 250 291.66666666666663 ${upperControlY} 500 250 Z`;
      }
      if (element.geometry === "lightningBolt") {
        const scale2 = 1e3 / 21600;
        return [
          [8472, 0],
          [12860, 6080],
          [11050, 6797],
          [16577, 12007],
          [14767, 12877],
          [21600, 21600],
          [10012, 14915],
          [12222, 13987],
          [5022, 9705],
          [7602, 8382],
          [0, 3890]
        ].map(([pointX, pointY], index) => (index ? "L" : "M") + pointX * scale2 + " " + pointY * scale2).join(" ") + " Z";
      }
      if (element.geometry === "triangle") return `M0 1000 L${((_Va = element.presetAdjustment) != null ? _Va : 0.5) * 1e3} 0 L1000 1000 Z`;
      if (element.geometry === "rtTriangle") return "M0 1000 L0 0 L1000 1000 Z";
      if (element.geometry === "diamond" || element.geometry === "flowChartDecision") return "M0 500 L500 0 L1000 500 L500 1000 Z";
      if (element.geometry === "flowChartOffpageConnector") return "M0 0 H1000 V800 L500 1000 L0 800 Z";
      if (element.geometry === "flowChartDocument") {
        const scale2 = 1e3 / 21600;
        return `M0 0 H1000 V${17322 * scale2} C${10800 * scale2} ${17322 * scale2} ${10800 * scale2} ${23922 * scale2} 0 ${20172 * scale2} Z`;
      }
      if (element.geometry === "flowChartManualInput") return "M0 200 L1000 0 V1000 H0 Z";
      if (element.geometry === "flowChartPunchedCard") return "M0 200 L200 0 H1000 V1000 H0 Z";
      if (element.geometry === "flowChartOfflineStorage") return "M0 0 H1000 L500 1000 Z M400 800 H600";
      if (element.geometry === "flowChartMagneticDisk") return "M0 166.6667 A500 166.6667 0 0 1 1000 166.6667 L1000 833.3333 A500 166.6667 0 0 1 0 833.3333 Z";
      if (element.geometry === "flowChartMagneticDrum") {
        const inset = 1e3 / 6;
        return `M${inset} 0 H833.3333333333334 A${inset} 500 0 0 1 833.3333333333334 1000 L${inset} 1000 A${inset} 500 0 0 1 ${inset} 0 Z`;
      }
      if (element.geometry === "flowChartMagneticTape") {
        const angle = Math.atan2(element.height, element.width);
        const x2 = 500 + 500 * Math.cos(angle), y2 = 500 + 500 * Math.sin(angle);
        return `M500 1000 A500 500 0 0 1 0 500 A500 500 0 0 1 500 0 A500 500 0 0 1 1000 500 A500 500 0 0 1 ${x2} ${y2} L1000 ${y2} L1000 1000 Z`;
      }
      if (element.geometry === "flowChartMultidocument") {
        const scale2 = 1e3 / 21600;
        const unit = (value) => value * scale2;
        return `M0 ${unit(20782)} C${unit(9298)} ${unit(23542)} ${unit(9298)} ${unit(18022)} ${unit(18595)} ${unit(18022)} L${unit(18595)} ${unit(3675)} H0 Z M${unit(1532)} ${unit(3675)} L${unit(1532)} ${unit(1815)} L${unit(2e4)} ${unit(1815)} L${unit(2e4)} ${unit(16252)} C${unit(19298)} ${unit(16252)} ${unit(18595)} ${unit(16352)} ${unit(18595)} ${unit(16352)} L${unit(18595)} ${unit(3675)} Z M${unit(2972)} ${unit(1815)} L${unit(2972)} 0 H1000 L1000 ${unit(14392)} C${unit(20800)} ${unit(14392)} ${unit(2e4)} ${unit(14467)} ${unit(2e4)} ${unit(14467)} L${unit(2e4)} ${unit(1815)} Z`;
      }
      if (element.geometry === "flowChartOnlineStorage") {
        const inset = 1e3 / 6;
        return `M${inset} 0 H1000 A${inset} 500 0 0 0 1000 1000 L${inset} 1000 A${inset} 500 0 0 1 ${inset} 0 Z`;
      }
      if (element.geometry === "flowChartPunchedTape") return "M0 100 A250 100 0 0 0 500 100 A250 100 0 0 1 1000 100 L1000 900 A250 100 0 0 0 500 900 A250 100 0 0 1 0 900 Z";
      if (element.geometry === "flowChartManualOperation") return "M0 0 H1000 L800 1000 L200 1000 Z";
      if (element.geometry === "flowChartInputOutput") return "M0 1000 L200 0 H1000 L800 1000 Z";
      if (element.geometry === "flowChartPreparation") return "M0 500 L200 0 H800 L1000 500 L800 1000 H200 Z";
      if (element.geometry === "flowChartTerminator") {
        const radiusX = 3475 / 21600 * 1e3;
        return `M${radiusX} 0 H839.1203703703704 A${radiusX} 500 0 0 1 1000 500 L${radiusX} 1000 A${radiusX} 500 0 0 1 0 500 Z`;
      }
      if (element.geometry === "flowChartPredefinedProcess") return "M0 0 H1000 V1000 H0 Z M125 0 V1000 M875 0 V1000";
      if (element.geometry === "flowChartInternalStorage") return "M0 0 H1000 V1000 H0 Z M125 0 V1000 M0 125 H1000";
      if (element.geometry === "flowChartSort") return "M0 500 L500 0 L1000 500 L500 1000 Z M0 500 H1000";
      if (element.geometry === "flowChartCollate") return "M0 0 H1000 L500 500 L1000 1000 H0 L500 500 Z";
      if (element.geometry === "flowChartExtract") return "M0 1000 L500 0 L1000 1000 Z";
      if (element.geometry === "flowChartMerge") return "M0 0 L1000 0 L500 1000 Z";
      if (element.geometry === "trapezoid") {
        const cut = element.width > 0 ? Math.min(element.width, element.height) * ((_Wa = element.presetAdjustment) != null ? _Wa : 0.25) / element.width * 1e3 : 0;
        return `M0 1000 L${cut} 0 L${1e3 - cut} 0 L1000 1000 Z`;
      }
      if (element.geometry === "octagon") {
        const shortSide = Math.min(element.width, element.height);
        const adjustment = (_Xa = element.presetAdjustment) != null ? _Xa : 0.29289;
        const cutX = element.width > 0 ? shortSide * adjustment / element.width * 1e3 : 0;
        const cutY = element.height > 0 ? shortSide * adjustment / element.height * 1e3 : 0;
        return `M0 ${cutY} L${cutX} 0 H${1e3 - cutX} L1000 ${cutY} V${1e3 - cutY} L${1e3 - cutX} 1000 H${cutX} L0 ${1e3 - cutY} Z`;
      }
      if (element.geometry === "pentagon") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const halfWidth = width * ((_Za = (_Ya = element.presetAdjustments) == null ? void 0 : _Ya.hf) != null ? _Za : 105146) / 2e5;
        const halfHeight = height * ((_$a = (__a = element.presetAdjustments) == null ? void 0 : __a.vf) != null ? _$a : 110557) / 2e5;
        const dx1 = halfWidth * Math.cos(Math.PI / 10), dx2 = halfWidth * Math.cos(17 * Math.PI / 10);
        const dy1 = halfHeight * Math.sin(Math.PI / 10), dy2 = halfHeight * Math.sin(17 * Math.PI / 10);
        const centerY = height * ((_bb = (_ab = element.presetAdjustments) == null ? void 0 : _ab.vf) != null ? _bb : 110557) / 2e5;
        const pointX = (value) => value / width * 1e3, pointY = (value) => value / height * 1e3;
        return `M${pointX(width / 2 - dx1)} ${pointY(centerY - dy1)} L500 0 L${pointX(width / 2 + dx1)} ${pointY(centerY - dy1)} L${pointX(width / 2 + dx2)} ${pointY(centerY - dy2)} L${pointX(width / 2 - dx2)} ${pointY(centerY - dy2)} Z`;
      }
      if (element.geometry === "heptagon") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const halfWidth = width / 2 * ((_db = (_cb = element.presetAdjustments) == null ? void 0 : _cb.hf) != null ? _db : 102572) / 1e5;
        const halfHeight = height / 2 * ((_fb = (_eb = element.presetAdjustments) == null ? void 0 : _eb.vf) != null ? _fb : 105210) / 1e5;
        const dx1 = halfWidth * 97493 / 1e5, dx2 = halfWidth * 78183 / 1e5, dx3 = halfWidth * 43388 / 1e5;
        const y1 = halfHeight - halfHeight * 62349 / 1e5;
        const y2 = halfHeight + halfHeight * 22252 / 1e5;
        const y3 = halfHeight + halfHeight * 90097 / 1e5;
        const x2 = (value) => value / width * 1e3, y4 = (value) => value / height * 1e3;
        return `M${x2(width / 2 - dx1)} ${y4(y2)} L${x2(width / 2 - dx2)} ${y4(y1)} L500 0 L${x2(width / 2 + dx2)} ${y4(y1)} L${x2(width / 2 + dx1)} ${y4(y2)} L${x2(width / 2 + dx3)} ${y4(y3)} L${x2(width / 2 - dx3)} ${y4(y3)} Z`;
      }
      if (element.geometry === "decagon") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const halfHeight = height / 2 * ((_hb = (_gb = element.presetAdjustments) == null ? void 0 : _gb.vf) != null ? _hb : 105146) / 1e5;
        const dx1 = width / 2 * Math.cos(Math.PI / 5), dx2 = width / 2 * Math.cos(2 * Math.PI / 5);
        const dy1 = halfHeight * Math.sin(2 * Math.PI / 5), dy2 = halfHeight * Math.sin(Math.PI / 5);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 500 L${x2(width / 2 - dx1)} ${y2(height / 2 - dy2)} L${x2(width / 2 - dx2)} ${y2(height / 2 - dy1)} L${x2(width / 2 + dx2)} ${y2(height / 2 - dy1)} L${x2(width / 2 + dx1)} ${y2(height / 2 - dy2)} L1000 500 L${x2(width / 2 + dx1)} ${y2(height / 2 + dy2)} L${x2(width / 2 + dx2)} ${y2(height / 2 + dy1)} L${x2(width / 2 - dx2)} ${y2(height / 2 + dy1)} L${x2(width / 2 - dx1)} ${y2(height / 2 + dy2)} Z`;
      }
      if (element.geometry === "star5") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const horizontalFactor = (_jb = (_ib = element.presetAdjustments) == null ? void 0 : _ib.hf) != null ? _jb : 105146, verticalFactor = (_lb = (_kb = element.presetAdjustments) == null ? void 0 : _kb.vf) != null ? _lb : 110557;
        const halfWidth = width / 2 * horizontalFactor / 1e5, halfHeight = height / 2 * verticalFactor / 1e5;
        const centerY = height / 2 * verticalFactor / 1e5;
        const innerScale = Math.max(0, Math.min(0.5, (_mb = element.presetAdjustment) != null ? _mb : 0.19098)) * 2;
        const innerHalfWidth = halfWidth * innerScale, innerHalfHeight = halfHeight * innerScale;
        const dx1 = halfWidth * Math.cos(Math.PI / 10), dx2 = halfWidth * Math.cos(17 * Math.PI / 10);
        const dy1 = halfHeight * Math.sin(Math.PI / 10), dy2 = halfHeight * Math.sin(17 * Math.PI / 10);
        const sdx1 = innerHalfWidth * Math.cos(19 * Math.PI / 10), sdx2 = innerHalfWidth * Math.cos(3 * Math.PI / 10);
        const sdy1 = innerHalfHeight * Math.sin(3 * Math.PI / 10), sdy2 = innerHalfHeight * Math.sin(19 * Math.PI / 10);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(width / 2 - dx1)} ${y2(centerY - dy1)} L${x2(width / 2 - sdx2)} ${y2(centerY - sdy1)} L500 0 L${x2(width / 2 + sdx2)} ${y2(centerY - sdy1)} L${x2(width / 2 + dx1)} ${y2(centerY - dy1)} L${x2(width / 2 + sdx1)} ${y2(centerY - sdy2)} L${x2(width / 2 + dx2)} ${y2(centerY - dy2)} L500 ${y2(centerY + innerHalfHeight)} L${x2(width / 2 - dx2)} ${y2(centerY - dy2)} L${x2(width / 2 - sdx1)} ${y2(centerY - sdy2)} Z`;
      }
      if (element.geometry === "star4") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const innerScale = Math.max(0, Math.min(0.5, (_nb = element.presetAdjustment) != null ? _nb : 0.125)) * 2;
        const innerHalfWidth = width / 2 * innerScale, innerHalfHeight = height / 2 * innerScale;
        const insetX = innerHalfWidth * Math.cos(Math.PI / 4), insetY = innerHalfHeight * Math.sin(Math.PI / 4);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 500 L${x2(width / 2 - insetX)} ${y2(height / 2 - insetY)} L500 0 L${x2(width / 2 + insetX)} ${y2(height / 2 - insetY)} L1000 500 L${x2(width / 2 + insetX)} ${y2(height / 2 + insetY)} L500 1000 L${x2(width / 2 - insetX)} ${y2(height / 2 + insetY)} Z`;
      }
      if (element.geometry === "star6") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const halfWidth = width / 2 * ((_pb = (_ob = element.presetAdjustments) == null ? void 0 : _ob.hf) != null ? _pb : 115470) / 1e5;
        const innerScale = Math.max(0, Math.min(0.5, (_qb = element.presetAdjustment) != null ? _qb : 0.28868)) * 2;
        const innerHalfWidth = halfWidth * innerScale, innerHalfHeight = height / 2 * innerScale;
        const dx = halfWidth * Math.cos(Math.PI / 6), dy = innerHalfHeight * Math.sin(Math.PI / 3);
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M${x2(width / 2 - dx)} ${y2(height / 4)} L${x2(width / 2 - innerHalfWidth / 2)} ${y2(height / 2 - dy)} L500 0 L${x2(width / 2 + innerHalfWidth / 2)} ${y2(height / 2 - dy)} L${x2(width / 2 + dx)} ${y2(height / 4)} L${x2(width / 2 + innerHalfWidth)} 500 L${x2(width / 2 + dx)} ${y2(height * 3 / 4)} L${x2(width / 2 + innerHalfWidth / 2)} ${y2(height / 2 + dy)} L500 1000 L${x2(width / 2 - innerHalfWidth / 2)} ${y2(height / 2 + dy)} L${x2(width / 2 - dx)} ${y2(height * 3 / 4)} L${x2(width / 2 - innerHalfWidth)} 500 Z`;
      }
      if (element.geometry === "star7") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const horizontalFactor = (_sb = (_rb = element.presetAdjustments) == null ? void 0 : _rb.hf) != null ? _sb : 102572, verticalFactor = (_ub = (_tb = element.presetAdjustments) == null ? void 0 : _tb.vf) != null ? _ub : 105210;
        const outerHalfWidth = width / 2 * horizontalFactor / 1e5, outerHalfHeight = height / 2 * verticalFactor / 1e5;
        const centerX = width / 2, centerY = height / 2 * verticalFactor / 1e5;
        const outerDx1 = outerHalfWidth * 0.97493, outerDx2 = outerHalfWidth * 0.78183, outerDx3 = outerHalfWidth * 0.43388;
        const outerDy1 = outerHalfHeight * 0.62349, outerDy2 = outerHalfHeight * 0.22252, outerDy3 = outerHalfHeight * 0.90097;
        const innerScale = Math.max(0, Math.min(0.5, (_vb = element.presetAdjustment) != null ? _vb : 0.34601)) * 2;
        const innerHalfWidth = outerHalfWidth * innerScale, innerHalfHeight = outerHalfHeight * innerScale;
        const innerDx1 = innerHalfWidth * 0.97493, innerDx2 = innerHalfWidth * 0.78183, innerDx3 = innerHalfWidth * 0.43388;
        const innerDy1 = innerHalfHeight * 0.90097, innerDy2 = innerHalfHeight * 0.22252, innerDy3 = innerHalfHeight * 0.62349;
        const x1 = centerX - outerDx1, x2 = centerX - outerDx2, x3 = centerX - outerDx3;
        const x4 = centerX + outerDx3, x5 = centerX + outerDx2, x6 = centerX + outerDx1;
        const y1 = centerY - outerDy1, y2 = centerY + outerDy2, y3 = centerY + outerDy3;
        const sx1 = centerX - innerDx1, sx2 = centerX - innerDx2, sx3 = centerX - innerDx3;
        const sx4 = centerX + innerDx3, sx5 = centerX + innerDx2, sx6 = centerX + innerDx1;
        const sy1 = centerY - innerDy1, sy2 = centerY - innerDy2, sy3 = centerY + innerDy3, sy4 = centerY + innerHalfHeight;
        const x7 = (value) => value / width * 1e3, y4 = (value) => value / height * 1e3;
        return [
          [x1, y2],
          [sx1, sy2],
          [x2, y1],
          [sx3, sy1],
          [centerX, 0],
          [sx4, sy1],
          [x5, y1],
          [sx6, sy2],
          [x6, y2],
          [sx5, sy3],
          [x4, y3],
          [centerX, sy4],
          [x3, y3],
          [sx2, sy3]
        ].map(([pointX, pointY], index) => (index ? "L" : "M") + x7(pointX) + " " + y4(pointY)).join(" ") + " Z";
      }
      if (element.geometry === "star8") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const innerScale = Math.max(0, Math.min(0.5, (_wb = element.presetAdjustment) != null ? _wb : 0.375)) * 2;
        const innerHalfWidth = width / 2 * innerScale, innerHalfHeight = height / 2 * innerScale;
        const outerDx = width / 2 * Math.cos(Math.PI / 4), outerDy = height / 2 * Math.sin(Math.PI / 4);
        const innerDx1 = innerHalfWidth * 0.92388, innerDx2 = innerHalfWidth * 0.38268;
        const innerDy1 = innerHalfHeight * 0.92388, innerDy2 = innerHalfHeight * 0.38268;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return `M0 500 L${x2(width / 2 - innerDx1)} ${y2(height / 2 - innerDy2)} L${x2(width / 2 - outerDx)} ${y2(height / 2 - outerDy)} L${x2(width / 2 - innerDx2)} ${y2(height / 2 - innerDy1)} L500 0 L${x2(width / 2 + innerDx2)} ${y2(height / 2 - innerDy1)} L${x2(width / 2 + outerDx)} ${y2(height / 2 - outerDy)} L${x2(width / 2 + innerDx1)} ${y2(height / 2 - innerDy2)} L1000 500 L${x2(width / 2 + innerDx1)} ${y2(height / 2 + innerDy2)} L${x2(width / 2 + outerDx)} ${y2(height / 2 + outerDy)} L${x2(width / 2 + innerDx2)} ${y2(height / 2 + innerDy1)} L500 1000 L${x2(width / 2 - innerDx2)} ${y2(height / 2 + innerDy1)} L${x2(width / 2 - outerDx)} ${y2(height / 2 + outerDy)} L${x2(width / 2 - innerDx1)} ${y2(height / 2 + innerDy2)} Z`;
      }
      if (element.geometry === "star10") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const outerHalfWidth = width / 2 * ((_yb = (_xb = element.presetAdjustments) == null ? void 0 : _xb.hf) != null ? _yb : 105146) / 1e5;
        const outerDx1 = outerHalfWidth * 0.95106, outerDx2 = outerHalfWidth * 0.58779;
        const outerDy1 = height / 2 * 0.80902, outerDy2 = height / 2 * 0.30902;
        const innerScale = Math.max(0, Math.min(0.5, (_zb = element.presetAdjustment) != null ? _zb : 0.42533)) * 2;
        const innerHalfWidth = outerHalfWidth * innerScale, innerHalfHeight = height / 2 * innerScale;
        const innerDx1 = innerHalfWidth * 0.80902, innerDx2 = innerHalfWidth * 0.30902;
        const innerDy1 = innerHalfHeight * 0.95106, innerDy2 = innerHalfHeight * 0.58779;
        const centerX = width / 2, centerY = height / 2;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return [
          [centerX - outerDx1, centerY - outerDy2],
          [centerX - innerDx1, centerY - innerDy2],
          [centerX - outerDx2, centerY - outerDy1],
          [centerX - innerDx2, centerY - innerDy1],
          [centerX, 0],
          [centerX + innerDx2, centerY - innerDy1],
          [centerX + outerDx2, centerY - outerDy1],
          [centerX + innerDx1, centerY - innerDy2],
          [centerX + outerDx1, centerY - outerDy2],
          [centerX + innerHalfWidth, centerY],
          [centerX + outerDx1, centerY + outerDy2],
          [centerX + innerDx1, centerY + innerDy2],
          [centerX + outerDx2, centerY + outerDy1],
          [centerX + innerDx2, centerY + innerDy1],
          [centerX, height],
          [centerX - innerDx2, centerY + innerDy1],
          [centerX - outerDx2, centerY + outerDy1],
          [centerX - innerDx1, centerY + innerDy2],
          [centerX - outerDx1, centerY + outerDy2],
          [centerX - innerHalfWidth, centerY]
        ].map(([pointX, pointY], index) => (index ? "L" : "M") + x2(pointX) + " " + y2(pointY)).join(" ") + " Z";
      }
      if (element.geometry === "star12") {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const centerX = width / 2, centerY = height / 2;
        const outerDx = centerX * Math.cos(Math.PI / 6), outerDy = centerY * Math.sin(Math.PI / 3);
        const innerScale = Math.max(0, Math.min(0.5, (_Ab = element.presetAdjustment) != null ? _Ab : 0.375)) * 2;
        const innerHalfWidth = centerX * innerScale, innerHalfHeight = centerY * innerScale;
        const innerDx1 = innerHalfWidth * Math.cos(Math.PI / 12), innerDx2 = innerHalfWidth * Math.cos(Math.PI / 4), innerDx3 = innerHalfWidth * Math.cos(5 * Math.PI / 12);
        const innerDy1 = innerHalfHeight * Math.sin(5 * Math.PI / 12), innerDy2 = innerHalfHeight * Math.sin(Math.PI / 4), innerDy3 = innerHalfHeight * Math.sin(Math.PI / 12);
        const x1 = centerX - outerDx, x3 = width * 3 / 4, x4 = centerX + outerDx;
        const y1 = centerY - outerDy, y3 = height * 3 / 4, y4 = centerY + outerDy;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        return [
          [0, centerY],
          [centerX - innerDx1, centerY - innerDy3],
          [x1, height / 4],
          [centerX - innerDx2, centerY - innerDy2],
          [width / 4, y1],
          [centerX - innerDx3, centerY - innerDy1],
          [centerX, 0],
          [centerX + innerDx3, centerY - innerDy1],
          [x3, y1],
          [centerX + innerDx2, centerY - innerDy2],
          [x4, height / 4],
          [centerX + innerDx1, centerY - innerDy3],
          [width, centerY],
          [centerX + innerDx1, centerY + innerDy3],
          [x4, y3],
          [centerX + innerDx2, centerY + innerDy2],
          [x3, y4],
          [centerX + innerDx3, centerY + innerDy1],
          [centerX, height],
          [centerX - innerDx3, centerY + innerDy1],
          [width / 4, y4],
          [centerX - innerDx2, centerY + innerDy2],
          [x1, y3],
          [centerX - innerDx1, centerY + innerDy3]
        ].map(([pointX, pointY], index) => (index ? "L" : "M") + x2(pointX) + " " + y2(pointY)).join(" ") + " Z";
      }
      if ([
        "star16",
        "star24",
        "star32"
      ].includes(element.geometry)) {
        const width = element.width, height = element.height;
        if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
        const pointCount = Number(element.geometry.slice(4));
        const innerScale = Math.max(0, Math.min(0.5, (_Bb = element.presetAdjustment) != null ? _Bb : 0.375)) * 2;
        const centerX = width / 2, centerY = height / 2;
        const guideFactor = (value) => Math.round(value * 1e5) / 1e5;
        const x2 = (value) => value / width * 1e3, y2 = (value) => value / height * 1e3;
        const points = [];
        for (let index = 0; index < pointCount * 2; index++) {
          const angle = Math.PI + index * Math.PI / pointCount;
          const radius = index % 2 === 0 ? 1 : innerScale;
          const usesExactTrig = element.geometry === "star24" && index % 2 === 0 || element.geometry === "star32" && index % 8 === 0;
          const factor = (value) => usesExactTrig ? value : guideFactor(value);
          points.push([centerX + centerX * radius * factor(Math.cos(angle)), centerY + centerY * radius * factor(Math.sin(angle))]);
        }
        return points.map(([pointX, pointY], index) => (index ? "L" : "M") + x2(pointX) + " " + y2(pointY)).join(" ") + " Z";
      }
      if (element.geometry === "dodecagon") {
        const scale2 = 1e3 / 21600;
        const x1 = 2894 * scale2, x2 = 7906 * scale2, x3 = 13694 * scale2, x4 = 18706 * scale2;
        return `M0 ${x2} L${x1} ${x1} L${x2} 0 H${x3} L${x4} ${x1} L1000 ${x2} V${x3} L${x4} ${x4} L${x3} 1000 H${x2} L${x1} ${x4} L0 ${x3} Z`;
      }
      if (element.geometry === "parallelogram" || element.geometry === "chevron" || element.geometry === "hexagon") {
        const adjustment = (_Cb = element.presetAdjustment) != null ? _Cb : element.geometry === "parallelogram" ? 0.25 : element.geometry === "hexagon" ? 0.25 : 0.5;
        const shortSide = Math.min(element.width, element.height);
        const cut = shortSide > 0 && element.width > 0 ? shortSide * adjustment / element.width * 1e3 : 0;
        if (element.geometry === "parallelogram") return `M0 1000 L${cut} 0 H1000 L${1e3 - cut} 1000 Z`;
        if (element.geometry === "chevron") return `M0 0 H${1e3 - cut} L1000 500 L${1e3 - cut} 1000 H0 L${cut} 500 Z`;
        return `M0 500 L${cut} ${0.211325 * 1e3} H${1e3 - cut} L1000 500 L${1e3 - cut} 788.675 H${cut} Z`;
      }
      if ([
        "ellipse",
        "flowChartConnector",
        "flowChartSummingJunction",
        "flowChartOr"
      ].includes(element.geometry)) {
        const ellipse = "M0 500 A500 500 0 1 0 1000 500 A500 500 0 1 0 0 500 Z";
        if (element.geometry === "flowChartSummingJunction") {
          const inset = (1 - Math.SQRT1_2) * 500;
          return `${ellipse} M${inset} ${inset} L${1e3 - inset} ${1e3 - inset} M${1e3 - inset} ${inset} L${inset} ${1e3 - inset}`;
        }
        if (element.geometry === "flowChartOr") return `${ellipse} M500 0 V1000 M0 500 H1000`;
        return ellipse;
      }
      if (element.geometry === "flowChartDelay") return "M0 0 H500 A500 500 0 0 1 500 1000 H0 Z";
      if (element.geometry === "flowChartDisplay") {
        const inset = 1e3 / 6;
        return `M0 500 L${inset} 0 H833.3333333333334 A${inset} 500 0 0 1 833.3333333333334 1000 L${inset} 1000 Z`;
      }
      if (element.geometry !== "roundRect" && element.geometry !== "flowChartAlternateProcess") return "M0 0 H1000 V1000 H0 Z";
      const [x, y] = roundRectRadii(element);
      if (!x || !y) return "M0 0 H1000 V1000 H0 Z";
      return `M0 ${y} A${x} ${y} 0 0 1 ${x} 0 H${1e3 - x} A${x} ${y} 0 0 1 1000 ${y} V${1e3 - y} A${x} ${y} 0 0 1 ${1e3 - x} 1000 H${x} A${x} ${y} 0 0 1 0 ${1e3 - y} Z`;
    }
    function customGeometryFill(element, path) {
      if (!path.fill || !animationOnOff(element, "fill.on")) return "none";
      const animatedFill = animationColor(element, "fillcolor");
      const patternFallback = (isPatternFillCss(element.fill) ? [...element.fill.matchAll(/#[\da-f]{6}\b|rgba?\([^)]*\)/gi)].map((match) => match[0]) : []).filter((color) => !/^rgba\([^)]*,\s*0(?:\.0+)?\s*\)$/i.test(color)).at(-1);
      const brightness = customPathBrightness(path.fillMode);
      if (!animatedFill && patternFallback) return applyColorOpacity(patternFallback, animationOpacity(element, "fill.opacity"), brightness);
      if (!animatedFill && /^(?:linear|radial)-gradient\(/i.test(element.fill)) return `url(#${customGradientId(element, "fill", path.fillMode)})`;
      return applyColorOpacity(animatedFill || element.fill, animationOpacity(element, "fill.opacity"), brightness);
    }
    function customPathBrightness(mode) {
      return mode ? {
        lighten: 0.4,
        lightenLess: 0.2,
        darken: -0.4,
        darkenLess: -0.2
      }[mode] : 0;
    }
    function isPatternFillCss(value) {
      return /repeating-linear-gradient\(/i.test(value) || /\) 0 0 \/ \d+(?:\.\d+)?px \d+(?:\.\d+)?px,/i.test(value);
    }
    function customGeometryStroke(element, path) {
      if (!path.stroke || !animationOnOff(element, "stroke.on")) return "none";
      const animatedStroke = animationColor(element, "stroke.color");
      const strokeWidth = animationStrokeWidth(element);
      if (path.stroke && strokeWidth > 0 && !animatedStroke && /^(?:linear|radial)-gradient\(/i.test(element.stroke)) return `url(#${customGradientId(element, "stroke")})`;
      const stroke = applyColorOpacity(animatedStroke || element.stroke, animationOpacity(element, "stroke.opacity"));
      return path.stroke && strokeWidth > 0 && stroke !== "transparent" && !stroke.includes("gradient") ? stroke : "none";
    }
    function customGradientId(element, kind, mode) {
      return `${customGradientScope}-${kind}-${element.id.replace(/[^a-zA-Z0-9_-]/g, "_")}${mode ? `-${mode}` : ""}`;
    }
    function customGradientFocusClipId(element, pathIndex) {
      return `${customGradientScope}-fill-focus-${element.id.replace(/[^a-zA-Z0-9_-]/g, "_")}-${pathIndex}`;
    }
    function splitGradientStops(value) {
      const stops = [];
      let depth = 0, start = 0;
      for (let index = 0; index < value.length; index++) if (value[index] === "(") depth++;
      else if (value[index] === ")") depth--;
      else if (value[index] === "," && depth === 0) {
        stops.push(value.slice(start, index).trim());
        start = index + 1;
      }
      stops.push(value.slice(start).trim());
      return stops;
    }
    const radialGradientBackgroundCache = /* @__PURE__ */ new Map();
    function radialGradientBackground(gradient) {
      const key = JSON.stringify(gradient);
      const cached = radialGradientBackgroundCache.get(key);
      if (cached) return cached;
      const stops = gradient.stops.map((stop) => `<stop offset="${stop.offset}" stop-color="${stop.color}" stop-opacity="${stop.opacity}"/>`).join("");
      const transform = gradient.gradientTransform ? ` gradientTransform="${gradient.gradientTransform}"` : "";
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" preserveAspectRatio="none"><defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="${gradient.cx}" cy="${gradient.cy}" r="${gradient.r}" fx="${gradient.fx}" fy="${gradient.fy}" fr="${gradient.fr}"${transform}${gradient.colorInterpolation ? ` color-interpolation="${gradient.colorInterpolation}"` : ""}>${stops}</radialGradient></defs><rect width="1000" height="1000" fill="url(#g)"/></svg>`;
      const background = `url("data:image/svg+xml,${encodeURIComponent(svg)}") center / 100% 100% no-repeat`;
      if (radialGradientBackgroundCache.size >= 128) radialGradientBackgroundCache.delete(radialGradientBackgroundCache.keys().next().value);
      radialGradientBackgroundCache.set(key, background);
      return background;
    }
    function customSvgGradient(element, kind, mode) {
      var _a, _b;
      const css = applyColorOpacity(kind === "fill" ? element.fill : element.stroke, 1, customPathBrightness(mode));
      const opacity = kind === "fill" ? animationOpacity(element, "fill.opacity") : 1;
      const layers = splitGradientStops(css);
      const linear = /^linear-gradient\(\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+))deg\s*,([\s\S]*)\)$/i.exec(layers.find((layer) => layer.startsWith("linear-gradient(")) || "");
      const radialLayer = layers.find((layer) => /^radial-gradient\(\s*(?:circle|ellipse)(?:\s+at\s+[-+\d.]+%\s+[-+\d.]+%)?\s*,/i.test(layer)) || "";
      const radial = /^radial-gradient\(\s*(circle|ellipse)(?:\s+at\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))%\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))%)?\s*,([\s\S]*)\)$/i.exec(radialLayer);
      if (!linear && !radial) return void 0;
      const stops = splitGradientStops(linear ? linear[2] : radial[4]).map((stop) => {
        const match = /^(.*)\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))%$/.exec(stop);
        if (!match) return void 0;
        const color = match[1].trim();
        const rgba = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*(?:,\s*([\d.]+)\s*)?\)$/i.exec(color);
        const hex = /^#([\da-f]{6})$/i.exec(color);
        if (color === "transparent") return {
          offset: `${Math.max(0, Math.min(Number(match[2]), 100))}%`,
          color: "#000000",
          opacity: 0
        };
        if (!rgba && !hex) return void 0;
        return {
          offset: `${Math.max(0, Math.min(Number(match[2]), 100))}%`,
          color: rgba ? `rgb(${rgba[1]}, ${rgba[2]}, ${rgba[3]})` : color,
          opacity: (rgba && rgba[4] !== void 0 ? Math.max(0, Math.min(Number(rgba[4]), 1)) : 1) * opacity
        };
      });
      if (stops.length < 2 || stops.some((stop) => !stop)) return void 0;
      if (radial) {
        const focusEllipseLayer = layers.find((layer) => /^radial-gradient\(\s*ellipse\s+closest-side\s*,[\s\S]*\)\s+left\s+[-+\d.]+%\s+top\s+[-+\d.]+%\s*\/\s*[-+\d.]+%\s+[-+\d.]+%\s+no-repeat$/i.test(layer));
        const focusEllipse = focusEllipseLayer && /^radial-gradient\(\s*ellipse\s+closest-side\s*,([\s\S]*)\)\s+left\s+([-+\d.]+)%\s+top\s+([-+\d.]+)%\s*\/\s*([-+\d.]+)%\s+([-+\d.]+)%\s+no-repeat$/i.exec(focusEllipseLayer);
        const focusRectLayer = layers.find((layer) => /^linear-gradient\([\s\S]*\)\s+left\s+[-+\d.]+%\s+top\s+[-+\d.]+%\s*\/\s*[-+\d.]+%\s+[-+\d.]+%\s+no-repeat$/i.test(layer));
        const focusRect = focusRectLayer && /^linear-gradient\(([\s\S]*)\)\s+left\s+([-+\d.]+)%\s+top\s+([-+\d.]+)%\s*\/\s*([-+\d.]+)%\s+([-+\d.]+)%\s+no-repeat$/i.exec(focusRectLayer);
        const focus = focusEllipse || focusRect;
        const focusColor = focusEllipse ? (_a = splitGradientStops(focusEllipse[1])[0]) == null ? void 0 : _a.replace(/\s+[-+\d.]+%$/, "").trim() : focusRect ? (_b = splitGradientStops(focusRect[1])[0]) == null ? void 0 : _b.trim() : void 0;
        const cx = radial[2] === void 0 ? 500 : Number(radial[2]) * 10;
        const cy = radial[3] === void 0 ? 500 : Number(radial[3]) * 10;
        const left = focus ? Number(focus[2]) : 0;
        const top = focus ? Number(focus[3]) : 0;
        const width = focus ? Number(focus[4]) : 0;
        const height = focus ? Number(focus[5]) : 0;
        let gradientTransform;
        let fx, fy, fr;
        let focusOverlay;
        if (element.width > 0 && element.height > 0) {
          const circleRadius = radial[1].toLowerCase() === "circle" ? Math.max(...[0, 1e3].flatMap((x) => [0, 1e3].map((y) => Math.hypot((x - cx) * element.width / 1e3, (y - cy) * element.height / 1e3)))) : void 0;
          const radiusX = circleRadius === void 0 ? Math.max(cx, 1e3 - cx) : circleRadius * 1e3 / element.width;
          const radiusY = circleRadius === void 0 ? Math.max(cy, 1e3 - cy) : circleRadius * 1e3 / element.height;
          const scaleX = radiusX / 500, scaleY = radiusY / 500;
          gradientTransform = `translate(${cx} ${cy}) scale(${scaleX} ${scaleY}) translate(${-cx} ${-cy})`;
          if (focusEllipse && radial[1].toLowerCase() === "circle" && circleRadius !== void 0) {
            const focusCx = (left + width / 2) * 10;
            const focusCy = (top + height / 2) * 10;
            const focusRadius = (width * element.width / 200 + height * element.height / 200) / 2;
            fx = cx + (focusCx - cx) / scaleX;
            fy = cy + (focusCy - cy) / scaleY;
            fr = focusRadius * 500 / circleRadius;
          } else if (focusEllipse) focusOverlay = {
            kind: "ellipse",
            cx: (left + width / 2) * 10,
            cy: (top + height / 2) * 10,
            rx: width * 5,
            ry: height * 5,
            color: focusColor || ""
          };
          else if (focusRect) focusOverlay = {
            kind: "rect",
            x: left * 10,
            y: top * 10,
            width: width * 10,
            height: height * 10,
            color: focusColor || ""
          };
        }
        return {
          id: customGradientId(element, kind, mode),
          kind: "radial",
          cx,
          cy,
          r: 500,
          fx,
          fy,
          fr,
          colorInterpolation: fr !== void 0 ? "linearRGB" : void 0,
          gradientTransform,
          stops,
          ...focusOverlay ? { focus: focusOverlay } : {}
        };
      }
      const radians = Number(linear[1]) * Math.PI / 180;
      const dx = Math.sin(radians), dy = -Math.cos(radians);
      const halfLength = (Math.abs(dx) + Math.abs(dy)) * 500;
      return {
        id: customGradientId(element, kind, mode),
        kind: "linear",
        x1: 500 - dx * halfLength,
        y1: 500 - dy * halfLength,
        x2: 500 + dx * halfLength,
        y2: 500 + dy * halfLength,
        stops
      };
    }
    function customGeometryGradients(element) {
      var _a;
      const modifiers = [...new Set((_a = element.customPaths) == null ? void 0 : _a.map((path) => path.fillMode).filter((mode) => mode))];
      return [...["fill", "stroke"].flatMap((kind) => {
        const gradient = customSvgGradient(element, kind);
        return gradient ? [gradient] : [];
      }), ...modifiers.flatMap((mode) => {
        const gradient = customSvgGradient(element, "fill", mode);
        return gradient ? [gradient] : [];
      })];
    }
    function customGeometryFillFocus(element, path) {
      if (!path.fill || !animationOnOff(element, "fill.on") || animationColor(element, "fillcolor")) return void 0;
      const gradient = customSvgGradient(element, "fill", path.fillMode);
      return (gradient == null ? void 0 : gradient.kind) === "radial" ? gradient.focus : void 0;
    }
    function customGeometryStrokeWidth(element) {
      return `${animationStrokeWidth(element) * scale.value}px`;
    }
    function customGeometryStrokeStyle(element) {
      const patterns = {
        dot: [1, 3],
        sysDot: [1, 1],
        dash: [4, 3],
        sysDash: [3, 1],
        lgDash: [8, 3],
        dashDot: [
          4,
          3,
          1,
          3
        ],
        sysDashDot: [
          3,
          1,
          1,
          1
        ],
        lgDashDot: [
          8,
          3,
          1,
          3
        ],
        lgDashDotDot: [
          8,
          3,
          1,
          3,
          1,
          3
        ],
        sysDashDotDot: [
          3,
          1,
          1,
          1,
          1,
          1
        ]
      };
      const strokeDash = animationStrokeDash(element);
      const pattern = Array.isArray(strokeDash) ? strokeDash : patterns[strokeDash || ""];
      const strokeWidth = animationStrokeWidth(element) * scale.value;
      const cap = element.strokeCap || "butt";
      const capLength = cap === "round" || !Array.isArray(strokeDash) && cap === "square" ? 1 : 0;
      return {
        ...pattern ? { strokeDasharray: pattern.map((length, index) => {
          const extension = Math.min(capLength, pattern[index % 2 ? index - 1 : index]);
          return `${(length + (index % 2 ? extension : -extension)) * strokeWidth}px`;
        }).join(" ") } : {},
        strokeOpacity: !animationColor(element, "stroke.color") && /gradient\(/i.test(element.stroke) ? animationOpacity(element, "stroke.opacity") : void 0,
        strokeLinecap: cap
      };
    }
    function cssLineStyle(pattern) {
      if (Array.isArray(pattern)) return "dashed";
      if (pattern === "dot" || pattern === "sysDot") return "dotted";
      return pattern && pattern !== "solid" ? "dashed" : "solid";
    }
    function lineStrokeColor(element) {
      if (!animationOnOff(element, "stroke.on")) return "transparent";
      return applyColorOpacity(animationColor(element, "stroke.color") || (element.stroke === "transparent" ? "#222" : element.stroke), animationOpacity(element, "stroke.opacity"));
    }
    function lineRenderStyle(element) {
      return {
        width: `${Math.max(element.width * scale.value, 1)}px`,
        height: `${Math.max(element.height * scale.value, 1)}px`,
        color: lineStrokeColor(element)
      };
    }
    function lineEndId(element, side) {
      return `pptx-line-${element.id}-${side}`;
    }
    function lineEndUrl(element, side) {
      const ending = animationLineEnd(element, side);
      return ending && ending.type !== "none" ? `url(#${lineEndId(element, side)})` : void 0;
    }
    function lineEndMarkerSize(element, side) {
      const ending = animationLineEnd(element, side);
      const size = {
        sm: 0.7,
        med: 1,
        lg: 1.4
      };
      return {
        width: 6 * size[ending.length],
        height: 6 * size[ending.width]
      };
    }
    function lineEndPath(type) {
      if (type === "stealth") return "M0 0L10 5L0 10L3.5 5Z";
      if (type === "diamond") return "M0 5L5 0L10 5L5 10Z";
      if (type === "oval") return "M5 0A5 5 0 1 1 5 10A5 5 0 1 1 5 0Z";
      if (type === "arrow" || type === "chevron") return "M0 0L10 5L0 10";
      if (type === "doubleChevron") return "M0 0L6 5L0 10M4 0L10 5L4 10";
      return "M0 0L10 5L0 10Z";
    }
    function lineEndUsesStroke(type) {
      return type === "arrow" || type === "chevron" || type === "doubleChevron";
    }
    function lineArrowType(value) {
      if (value === "block") return "triangle";
      if (value === "classic") return "stealth";
      if (value === "open") return "arrow";
      if (value === "doublechevron") return "doubleChevron";
      return value;
    }
    function lineArrowSize(value) {
      if (value === "narrow" || value === "short") return "sm";
      if (value === "wide" || value === "long") return "lg";
      return "med";
    }
    function animationLineEnd(element, side) {
      var _a, _b, _c;
      const property = side === "head" ? "stroke.startArrow" : "stroke.endArrow";
      const baseline = side === "head" ? element.lineHeadEnd : element.lineTailEnd;
      let type = (baseline == null ? void 0 : baseline.type) || "none";
      for (const action of ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => item.effect === "lineArrow" && item.lineArrowProperty === property)) || []) {
        const progress = animationProgress(action);
        if (progress === -2) {
          type = (baseline == null ? void 0 : baseline.type) || "none";
          continue;
        }
        if (progress === null || progress < 0) continue;
        if ((_b = action.lineArrowKeyframes) == null ? void 0 : _b.length) for (const frame of action.lineArrowKeyframes) {
          if (frame.offset > progress) break;
          type = lineArrowType(frame.value);
        }
        else if (action.lineArrowFrom !== void 0 && progress < 1) type = lineArrowType(action.lineArrowFrom);
        else if (progress >= 1) {
          const value = (_c = action.lineArrowTo) != null ? _c : action.lineArrowFrom;
          if (value !== void 0) type = lineArrowType(value);
        }
      }
      const baselineWidth = (baseline == null ? void 0 : baseline.width) || "med";
      const baselineLength = (baseline == null ? void 0 : baseline.length) || "med";
      const animatedSize = (property2, initial) => {
        var _a2, _b2, _c2;
        let value = initial;
        for (const action of ((_a2 = animationsByTarget.value.get(element.id)) == null ? void 0 : _a2.filter((item) => item.effect === "lineArrowSize" && item.lineArrowSizeProperty === property2)) || []) {
          const progress = animationProgress(action);
          if (progress === -2) {
            value = initial;
            continue;
          }
          if (progress === null || progress < 0) continue;
          if ((_b2 = action.lineArrowSizeKeyframes) == null ? void 0 : _b2.length) for (const frame of action.lineArrowSizeKeyframes) {
            if (frame.offset > progress) break;
            value = lineArrowSize(frame.value);
          }
          else if (action.lineArrowSizeFrom !== void 0 && progress < 1) value = lineArrowSize(action.lineArrowSizeFrom);
          else if (progress >= 1) {
            const finalValue = (_c2 = action.lineArrowSizeTo) != null ? _c2 : action.lineArrowSizeFrom;
            if (finalValue !== void 0) value = lineArrowSize(finalValue);
          }
        }
        return value;
      };
      const widthProperty = side === "head" ? "stroke.startArrowWidth" : "stroke.endArrowWidth";
      const lengthProperty = side === "head" ? "stroke.startArrowLength" : "stroke.endArrowLength";
      return {
        type,
        width: animatedSize(widthProperty, baselineWidth),
        length: animatedSize(lengthProperty, baselineLength)
      };
    }
    function customGeometryStyle(element) {
      var _a;
      const outerShadows = (((_a = element.shadows) == null ? void 0 : _a.length) ? element.shadows : element.shadow ? [element.shadow] : []).filter((shadow) => !shadow.inset);
      if (!outerShadows.length || !animationOnOff(element, "shadow.on")) return {};
      return { filter: outerShadows.map((shadow) => `drop-shadow(${shadowCss(shadow, applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity")))})`).join(" ") };
    }
    function shadowCss(shadow, color) {
      const angle = (shadow.angle || 0) * Math.PI / 180;
      return [
        ...shadow.inset ? ["inset"] : [],
        Math.cos(angle) * shadow.distance * scale.value + "px",
        Math.sin(angle) * shadow.distance * scale.value + "px",
        shadow.blur * scale.value + "px",
        color
      ].join(" ");
    }
    function pictureInnerShadowStyle(element) {
      var _a;
      const innerShadows = (((_a = element.shadows) == null ? void 0 : _a.length) ? element.shadows : element.shadow ? [element.shadow] : []).filter((shadow) => shadow.inset);
      if (!innerShadows.length || !animationOnOff(element, "shadow.on")) return void 0;
      return { boxShadow: innerShadows.map((shadow) => shadowCss(shadow, applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity")))).join(", ") };
    }
    function imageStyle(element) {
      var _a;
      const hasCropAnimation = (_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.some((item) => item.effect === "imageCrop");
      const crop = element.imageCrop || hasCropAnimation ? imageCrop(element) : void 0;
      const grayscale = element.kind === "picture" && imageGrayscale(element);
      const filters = [imageToneEnabled(element) ? `url(#${imageToneFilterId(element)})` : void 0, grayscale ? "grayscale(1)" : void 0].filter((filter2) => Boolean(filter2));
      const filter = filters.length ? filters.join(" ") : void 0;
      if (!crop) return {
        width: "100%",
        height: "100%",
        filter
      };
      const width = 1 - crop.left - crop.right;
      const height = 1 - crop.top - crop.bottom;
      return {
        width: `${100 / width}%`,
        height: `${100 / height}%`,
        left: `${-crop.left / width * 100}%`,
        top: `${-crop.top / height * 100}%`,
        filter
      };
    }
    function imageCrop(element) {
      var _a, _b, _c, _d, _e, _f;
      const base = element.imageCrop || {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0
      };
      const crop = { ...base };
      for (const action of ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => item.effect === "imageCrop")) || []) {
        const property = action.imageCropProperty;
        if (!property || action.imageCropTo === void 0 && !((_b = action.imageCropKeyframes) == null ? void 0 : _b.length)) continue;
        const progress = animationProgress(action);
        if (progress === -2) {
          crop[property] = base[property];
          continue;
        }
        if (progress === null || progress < 0) continue;
        const value = ((_c = action.imageCropKeyframes) == null ? void 0 : _c.length) ? animationKeyframeValue(action.imageCropKeyframes, action.imageCropKeyframeMode, progress) : ((_d = action.imageCropFrom) != null ? _d : crop[property]) + (((_e = action.imageCropTo) != null ? _e : 0) - ((_f = action.imageCropFrom) != null ? _f : crop[property])) * progress;
        if (value >= 0 && value < 1 && value + crop[property === "left" ? "right" : property === "right" ? "left" : property === "top" ? "bottom" : "top"] < 1) crop[property] = value;
      }
      return crop;
    }
    function imageGrayscale(element) {
      var _a, _b, _c, _d;
      let grayscale = element.imageGrayscale === true;
      for (const action of ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => item.effect === "grayscale")) || []) {
        const progress = animationProgress(action);
        if (progress === null || progress < 0) continue;
        let value = (_b = action.grayscaleFrom) != null ? _b : grayscale;
        if ((_c = action.grayscaleKeyframes) == null ? void 0 : _c.length) for (const frame of action.grayscaleKeyframes) {
          if (frame.offset > progress) break;
          value = frame.value;
        }
        else if (progress >= 1) value = (_d = action.grayscaleTo) != null ? _d : value;
        grayscale = value;
      }
      return grayscale;
    }
    function imageTone(element) {
      var _a, _b, _c, _d, _e;
      const tone = {
        gain: 1,
        blacklevel: 0,
        gamma: 1
      };
      for (const action of ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => item.effect === "imageTone")) || []) {
        const property = action.imageToneProperty;
        if (!property || action.imageToneTo === void 0 && !((_b = action.imageToneKeyframes) == null ? void 0 : _b.length)) continue;
        const progress = animationProgress(action);
        if (progress === -2) {
          tone[property] = property === "blacklevel" ? 0 : 1;
          continue;
        }
        if (progress === null || progress < 0) continue;
        const from = (_c = action.imageToneFrom) != null ? _c : tone[property];
        const value = ((_d = action.imageToneKeyframes) == null ? void 0 : _d.length) ? animationKeyframeValue(action.imageToneKeyframes, action.imageToneKeyframeMode, progress) : from + (((_e = action.imageToneTo) != null ? _e : from) - from) * progress;
        if (Number.isFinite(value)) tone[property] = value;
      }
      return tone;
    }
    function imageToneEnabled(element) {
      var _a;
      return element.kind === "picture" && Boolean((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.some((item) => item.effect === "imageTone"));
    }
    function imageToneFilterId(element) {
      return `${customGradientScope}-image-tone-${element.id}`.replace(/[^a-zA-Z0-9_-]/g, "_");
    }
    function mediaPlaybackRange(media, element) {
      const trim = element.mediaTrim;
      if (!trim) return {
        start: 0,
        end: media.duration,
        valid: true
      };
      const start = trim.startMs / 1e3;
      const end = media.duration - trim.endMs / 1e3;
      if (start < end) return {
        start,
        end,
        valid: true
      };
      if (!invalidMediaTrimIds.value.includes(element.id)) invalidMediaTrimIds.value = [...invalidMediaTrimIds.value, element.id];
      return {
        start: 0,
        end: media.duration,
        valid: false
      };
    }
    function finishTimedMedia(element, refresh = true) {
      var _a;
      const endedAt = performance.now();
      let changed = false;
      const media = Array.from(((_a = stage.value) == null ? void 0 : _a.querySelectorAll("[data-media-id]")) || []).find((item) => item.dataset.mediaId === element.id);
      if (media) {
        delete media.dataset.mediaActionEndAt;
        mediaBookmarkPlaybackStarted.delete(media);
        lastMediaBookmarkTimes.delete(media);
      }
      timelineActions.value.forEach((action, index) => {
        if (action.effect !== "media" || action.targetId !== element.id || !action.mediaActionKey || !executedMediaActions.has(index)) return;
        completedMediaActions.set(action.mediaActionKey, endedAt);
        const timer = mediaActionEndTimers.get(action.mediaActionKey);
        if (timer) clearTimeout(timer);
        mediaActionEndTimers.delete(action.mediaActionKey);
        changed = true;
      });
      if (changed && refresh) refreshTimeline();
    }
    function runTimedMediaActions() {
      timelineActions.value.forEach((action, index) => {
        var _a, _b, _c, _d, _e, _f;
        if (action.effect !== "media" || !action.mediaCommand || executedMediaActions.has(index)) return;
        const startedAt = timelineStart(action);
        const stoppedAt = timelineEnd(action);
        if (stoppedAt !== void 0 && startedAt !== void 0 && startedAt + action.delayMs >= stoppedAt) return;
        if (startedAt === void 0 || startedAt === Number.NEGATIVE_INFINITY || animationClock.value < startedAt + action.delayMs) return;
        executedMediaActions.add(index);
        if (action.mediaCommand === "stopAudio") {
          emit("stop-all-audio");
          for (const element2 of props.slide.elements) {
            if (element2.kind !== "audio") continue;
            const audio = Array.from(((_a = stage.value) == null ? void 0 : _a.querySelectorAll("[data-media-id]")) || []).find((item) => item.dataset.mediaId === element2.id);
            if (!audio) continue;
            if (audio.readyState > 0) audio.currentTime = (((_b = element2.mediaTrim) == null ? void 0 : _b.startMs) || 0) / 1e3;
            finishTimedMedia(element2, false);
          }
          runSlideTriggeredAnimations("onStopAudio");
          refreshTimeline();
          return;
        }
        const element = props.slide.elements.find((item) => item.id === action.targetId);
        const media = Array.from(((_c = stage.value) == null ? void 0 : _c.querySelectorAll("[data-media-id]")) || []).find((item) => item.dataset.mediaId === action.targetId);
        if (!element || !media) return;
        if (action.mediaSlideCount) media.dataset.mediaSlideCount = String(action.mediaSlideCount);
        if (action.mediaVolume !== void 0) media.volume = action.mediaVolume;
        if (action.mediaMuted !== void 0) media.muted = action.mediaMuted;
        const seek = (seconds) => {
          const apply = () => {
            if (media.readyState === 0 || !Number.isFinite(media.duration)) return;
            const range = mediaPlaybackRange(media, element);
            media.currentTime = range.valid ? Math.max(range.start, Math.min(seconds, range.end)) : Math.max(0, Math.min(seconds, media.duration));
          };
          if (media.readyState === 0 || !Number.isFinite(media.duration)) media.addEventListener("loadedmetadata", apply, { once: true });
          else apply();
        };
        const play = () => {
          media.play().catch((error) => {
            if (error instanceof DOMException && error.name === "NotAllowedError" && !blockedMediaPlaybackIds.value.includes(element.id)) blockedMediaPlaybackIds.value = [...blockedMediaPlaybackIds.value, element.id];
          });
        };
        if (action.mediaCommand === "play") {
          if (action.mediaStartSeconds !== void 0) seek(action.mediaStartSeconds);
          else if (media.currentTime < (((_d = element.mediaTrim) == null ? void 0 : _d.startMs) || 0) / 1e3) seek((((_e = element.mediaTrim) == null ? void 0 : _e.startMs) || 0) / 1e3);
          play();
          if (action.mediaActionKey && action.mediaDurationMs !== void 0) {
            media.dataset.mediaActionEndAt = String(performance.now() + action.mediaDurationMs);
            mediaActionEndTimers.set(action.mediaActionKey, setTimeout(() => {
              delete media.dataset.mediaActionEndAt;
              media.pause();
              finishTimedMedia(element);
            }, action.mediaDurationMs));
          }
        } else if (action.mediaCommand === "pause") media.pause();
        else if (action.mediaCommand === "togglePause") {
          if (media.paused) play();
          else media.pause();
        } else {
          media.pause();
          seek((((_f = element.mediaTrim) == null ? void 0 : _f.startMs) || 0) / 1e3);
          finishTimedMedia(element, false);
        }
      });
    }
    function easedProgress(action, progress) {
      return timingProgress(action, progress);
    }
    function finalAnimationProgress(action) {
      if (action.effectProgress !== void 0) return action.effectProgress;
      if (action.durationMs === 0) return action.reversePlayback ? 0 : 1;
      const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1);
      const totalDuration = animationDuration(action);
      if (totalDuration === 0) return 0;
      if (!Number.isFinite(totalDuration)) return 0;
      const remainder = totalDuration % cycleDuration;
      let progress;
      if (!action.autoReverse) progress = remainder === 0 ? 1 : easedProgress(action, remainder / action.durationMs);
      else if (remainder === 0) progress = 0;
      else if (remainder <= action.durationMs) progress = easedProgress(action, remainder / action.durationMs);
      else progress = 1 - easedProgress(action, (remainder - action.durationMs) / action.durationMs);
      return action.reversePlayback ? 1 - progress : progress;
    }
    function animationFillExpired(action, startedAt, iterationIndex = 0) {
      if (action.fillMode === "remove") {
        if (startedAt === void 0) return action.triggerSequence === void 0 && !action.automatic && action.step < props.animationStep;
        if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return true;
        const stoppedAt = timelineEnd(action);
        if (stoppedAt !== void 0 && animationClock.value >= stoppedAt) return true;
        const stagger = action.iteration ? action.iteration.intervalMs * iterationIndex : 0;
        return action.timingWarp ? timingWarpChildElapsed(action, startedAt, stagger) >= animationDuration(action) : timelineNow(action) >= startedAt + action.delayMs + stagger + animationDuration(action);
      }
      if (action.fillMode !== "freeze" || startedAt === void 0) return false;
      const actionStart = startedAt + action.delayMs;
      return (animationsByTarget.value.get(action.targetId) || []).some((sibling) => {
        var _a, _b;
        if (sibling.sequenceKey !== action.sequenceKey || ((_a = sibling.sequenceOrder) != null ? _a : -1) <= ((_b = action.sequenceOrder) != null ? _b : -1)) return false;
        const siblingStart = timelineStart(sibling);
        if (siblingStart === void 0) return false;
        const start = siblingStart + sibling.delayMs;
        return start >= actionStart && start <= timelineNow(action);
      });
    }
    function timingWarpChildElapsed(action, startedAt, stagger = 0) {
      const warp = action.timingWarp;
      const elapsedWallMs = timelineNow(action) - startedAt - warp.groupOffsetMs;
      if (elapsedWallMs < 0) return -1;
      const elapsedPlaybackMs = Math.min(elapsedWallMs, warp.groupPlaybackDurationMs) * warp.groupSpeed;
      const cycleDurationMs = warp.groupDurationMs * (warp.groupAutoReverse ? 2 : 1);
      let cyclePositionMs = elapsedPlaybackMs % cycleDurationMs;
      if (elapsedWallMs >= warp.groupPlaybackDurationMs && (cyclePositionMs < 0.01 || cycleDurationMs - cyclePositionMs < 0.01)) cyclePositionMs = cycleDurationMs;
      const reverseLeg = warp.groupAutoReverse && cyclePositionMs > warp.groupDurationMs;
      const legProgress = reverseLeg ? (cyclePositionMs - warp.groupDurationMs) / warp.groupDurationMs : cyclePositionMs / warp.groupDurationMs;
      return timingProgress(warp, warp.groupReversePlayback !== reverseLeg ? 1 - legProgress : legProgress) * warp.groupDurationMs - warp.childOffsetMs - stagger;
    }
    function animationProgress(action, iterationIndex) {
      if (props.thumbnail && !props.playbackPreview) return null;
      const now = timelineNow(action);
      if (action.iteration && (iterationIndex === void 0 || iterationIndex < 0)) return null;
      const startedAt = timelineStart(action);
      const stagger = action.iteration ? action.iteration.intervalMs * (iterationIndex || 0) : 0;
      if (animationFillExpired(action, startedAt, iterationIndex || 0)) return -2;
      if (startedAt === void 0) return action.triggerSequence === void 0 && !action.automatic && action.step < props.animationStep ? finalAnimationProgress(action) : null;
      if (startedAt === Infinity) return -1;
      if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return finalAnimationProgress(action);
      const elapsed = action.timingWarp ? timingWarpChildElapsed(action, startedAt, stagger) : now - startedAt - action.delayMs - stagger;
      if (elapsed < 0) return -1;
      if (action.effectProgress !== void 0) return action.effectProgress;
      if (action.durationMs === 0) return action.reversePlayback ? 0 : 1;
      const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1);
      const totalDuration = animationDuration(action);
      if (totalDuration === 0) return 0;
      if (elapsed >= totalDuration) return finalAnimationProgress(action);
      const position = elapsed % cycleDuration;
      const progress = action.autoReverse && position > action.durationMs ? 1 - easedProgress(action, (position - action.durationMs) / action.durationMs) : easedProgress(action, Math.min(position / action.durationMs, 1));
      return action.reversePlayback ? 1 - progress : progress;
    }
    function completedAnimationRepeats(action, iterationIndex) {
      const cycleDuration = action.durationMs * (action.autoReverse ? 2 : 1);
      if (!(action.rotationAccumulate || action.scaleAccumulate || action.motionAccumulate) || action.autoReverse || cycleDuration <= 0) return 0;
      const totalDuration = animationDuration(action);
      const lastRepeat = Number.isFinite(totalDuration) ? Math.max(0, Math.ceil(totalDuration / cycleDuration) - 1) : Infinity;
      const startedAt = timelineStart(action);
      if (startedAt === void 0) return action.triggerSequence === void 0 && !action.automatic && action.step < props.animationStep ? lastRepeat : 0;
      if (startedAt === Infinity) return 0;
      if (startedAt === Number.NEGATIVE_INFINITY || reduceMotion.value) return Number.isFinite(lastRepeat) ? lastRepeat : 0;
      const stagger = action.iteration ? action.iteration.intervalMs * (iterationIndex || 0) : 0;
      const elapsed = action.timingWarp ? timingWarpChildElapsed(action, startedAt, stagger) : timelineNow(action) - startedAt - action.delayMs - stagger;
      return Math.min(lastRepeat, Math.max(0, Math.floor(elapsed / cycleDuration)));
    }
    function motionAccumulationOffset(action, iterationIndex) {
      if (!action.motionAccumulate) return [0, 0];
      const points = action.motionPathSamples || action.motionPath;
      const start = points == null ? void 0 : points[0];
      const end = points == null ? void 0 : points[points.length - 1];
      if (!start || !end) return [0, 0];
      const repeats = completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1);
      return [(end.x - start.x) * repeats, (end.y - start.y) * repeats];
    }
    function animationScale(element) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "scale" && !action.paragraphRange && !action.characterRange && !action.iteration)) || [];
      let value = [1, 1];
      for (const action of actions) {
        const progress = animationProgress(action);
        if (progress === -2) {
          value = [1, 1];
          continue;
        }
        if (progress === null || progress < 0 || !action.scaleFrom || !action.scaleTo) continue;
        value = animationScaleValue(action, progress);
      }
      return value;
    }
    function animationScaleValue(action, progress, iterationIndex) {
      const [fromX, fromY] = action.scaleFrom;
      const [toX, toY] = action.scaleTo;
      const completedRepeats = action.scaleAccumulate ? completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1) : 0;
      return [fromX + (toX - fromX) * (progress + completedRepeats), fromY + (toY - fromY) * (progress + completedRepeats)];
    }
    function animationRotation(element) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "rotation" && !action.paragraphRange && !action.characterRange && !action.iteration)) || [];
      let value = element.rotation;
      for (const action of actions) {
        const progress = animationProgress(action);
        if (progress === -2) {
          value = element.rotation;
          continue;
        }
        if (progress === null || progress < 0 || action.rotationFrom === void 0 || action.rotationTo === void 0) continue;
        const delta = action.rotationTo - action.rotationFrom;
        const repeatOffset = delta * (action.reversePlayback ? -1 : 1) * completedAnimationRepeats(action);
        const rotation = action.rotationFrom + delta * progress + repeatOffset;
        value = action.rotationRelative || action.rotationAdditive ? value + rotation : rotation;
      }
      return value;
    }
    function animationKeyframeValue(keyframes, mode, progress) {
      if (mode === "discrete") {
        let frame = keyframes[0];
        for (const candidate of keyframes) {
          if (candidate.offset > progress) break;
          frame = candidate;
        }
        return frame.value;
      }
      const nextIndex = keyframes.findIndex((frame) => frame.offset >= progress);
      if (nextIndex === 0) return keyframes[0].value;
      if (nextIndex < 0) return keyframes.at(-1).value;
      const from = keyframes[nextIndex - 1], to = keyframes[nextIndex];
      const segmentProgress = (progress - from.offset) / (to.offset - from.offset);
      return from.value + (to.value - from.value) * segmentProgress;
    }
    function animationOpacity(element, property = "style.opacity", paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        if (action.effect !== "opacity" || action.opacityProperty !== property) return false;
        const hasTextRange = Boolean(action.paragraphRange || action.characterRange || action.iteration);
        if (paragraphIndex === void 0) return !hasTextRange;
        const iterationIndex = characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd);
        return hasTextRange && (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart !== void 0 && characterEnd !== void 0 && characterStart < action.characterRange.end && characterEnd > action.characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      let opacity = 1;
      for (const action of actions) {
        const progress = animationProgress(action, characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd));
        if (progress === -2) {
          opacity = 1;
          continue;
        }
        if (progress === null || progress < 0 || action.opacityFrom === void 0 || action.opacityTo === void 0) continue;
        const formulaSamples = action.opacityFormulaSamples;
        const keyframes = action.opacityKeyframes;
        if (formulaSamples == null ? void 0 : formulaSamples.length) opacity = animationKeyframeValue(formulaSamples, "lin", progress);
        else if (keyframes == null ? void 0 : keyframes.length) opacity = animationKeyframeValue(keyframes, action.opacityKeyframeMode, progress);
        else opacity = action.opacityFrom + (action.opacityTo - action.opacityFrom) * progress;
      }
      return opacity;
    }
    function animationOnOff(element, property) {
      var _a, _b, _c, _d;
      let enabled = true;
      for (const action of ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => item.effect === "onOff" && item.onOffProperty === property)) || []) {
        const progress = animationProgress(action);
        if (progress === -2) {
          enabled = true;
          continue;
        }
        if (progress === null || progress < 0) continue;
        if ((_b = action.onOffKeyframes) == null ? void 0 : _b.length) for (const frame of action.onOffKeyframes) {
          if (frame.offset > progress) break;
          enabled = frame.value;
        }
        else if (action.onOffFrom !== void 0 && progress < 1) enabled = action.onOffFrom;
        else if (progress >= 1) enabled = (_d = (_c = action.onOffTo) != null ? _c : action.onOffFrom) != null ? _d : enabled;
      }
      return enabled;
    }
    function animationStrokeDash(element) {
      var _a, _b, _c, _d;
      let strokeDash = element.strokeDash;
      for (const action of ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => item.effect === "strokeDash")) || []) {
        const progress = animationProgress(action);
        if (progress === -2) {
          strokeDash = element.strokeDash;
          continue;
        }
        if (progress === null || progress < 0) continue;
        if ((_b = action.strokeDashKeyframes) == null ? void 0 : _b.length) for (const frame of action.strokeDashKeyframes) {
          if (frame.offset > progress) break;
          strokeDash = frame.value;
        }
        else if (action.strokeDashFrom !== void 0 && progress < 1) strokeDash = action.strokeDashFrom;
        else if (progress >= 1) strokeDash = (_d = (_c = action.strokeDashTo) != null ? _c : action.strokeDashFrom) != null ? _d : strokeDash;
      }
      return strokeDash;
    }
    function animationStrokeWidth(element) {
      var _a, _b, _c;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "strokeWidth")) || [];
      let strokeWidth = element.strokeWidth;
      for (const action of actions) {
        const progress = animationProgress(action);
        if (progress === -2) {
          strokeWidth = element.strokeWidth;
          continue;
        }
        if (progress === null || progress < 0 || action.strokeWidthFrom === void 0 || action.strokeWidthTo === void 0) continue;
        strokeWidth = ((_b = action.strokeWidthFormulaSamples) == null ? void 0 : _b.length) ? animationKeyframeValue(action.strokeWidthFormulaSamples, "lin", progress) : ((_c = action.strokeWidthKeyframes) == null ? void 0 : _c.length) ? animationKeyframeValue(action.strokeWidthKeyframes, action.strokeWidthKeyframeMode, progress) : action.strokeWidthFrom + (action.strokeWidthTo - action.strokeWidthFrom) * progress;
      }
      if (strokeWidth === 0 && (animationColor(element, "stroke.color") || element.stroke) !== "transparent" && scale.value > 0) return 1 / (scale.value * pixelRatio.value);
      return strokeWidth;
    }
    function animationColor(element, property, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        if (action.effect !== "color" || action.colorProperty !== property) return false;
        const paragraphRange = action.paragraphRange;
        const characterRange = action.characterRange;
        const iterationIndex = characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd);
        return (!paragraphRange || paragraphIndex !== void 0 && paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart !== void 0 && characterEnd !== void 0 && characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      let color;
      for (const action of actions) {
        const progress = animationProgress(action, characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd));
        if (progress === -2) {
          color = void 0;
          continue;
        }
        const from = action.colorFrom, to = action.colorTo;
        if (progress === null || progress < 0 || !from || !to) continue;
        let channels;
        if (action.colorSpace === "hsl") {
          const hueDelta = ((to[0] - from[0]) % 1 + 1) % 1;
          const huePath = action.colorDirection === "ccw" && hueDelta ? hueDelta - 1 : hueDelta;
          const [r, g, b] = hslToRgb(from[0] + huePath * progress, from[1] + (to[1] - from[1]) * progress, from[2] + (to[2] - from[2]) * progress);
          channels = [
            r,
            g,
            b,
            from[3] + (to[3] - from[3]) * progress
          ];
        } else channels = [
          0,
          1,
          2,
          3
        ].map((index) => from[index] + (to[index] - from[index]) * progress);
        color = `rgba(${channels.slice(0, 3).map((channel) => Math.round(Math.max(0, Math.min(channel, 255)))).join(", ")}, ${Math.max(0, Math.min(channels[3], 1))})`;
      }
      return color;
    }
    function textColorAnimationStyle(element, paragraphIndex, characterStart, characterEnd) {
      const color = animationColor(element, "style.color", paragraphIndex, characterStart, characterEnd);
      if (!color) return {};
      return { color: animatedRunColor(color, animationOpacity(element, "fill.opacity", paragraphIndex, characterStart, characterEnd)) };
    }
    function runHasTextPattern(element, run, paragraphIndex, characterStart, characterEnd) {
      return Boolean(run.fillPattern) && !animationColor(element, "style.color", paragraphIndex, characterStart, characterEnd);
    }
    function textRangeOpacityAnimationStyle(element, paragraphIndex, characterStart, characterEnd) {
      const opacity = animationOpacity(element, "style.opacity", paragraphIndex, characterStart, characterEnd);
      return opacity === 1 ? {} : { opacity };
    }
    function textRunAnimationStyle(element, paragraphIndex, characterStart, characterEnd, run, fontScale = ((_a) => (_a = element.textFontScale) != null ? _a : 1)(), script = "latin", separateDecorations = false) {
      var _a2, _b, _c, _d, _e, _f, _g, _h;
      const actions = ((_a2 = animationsByTarget.value.get(element.id)) == null ? void 0 : _a2.filter((action) => {
        if (action.effect !== "fontSize") return false;
        const paragraphRange = action.paragraphRange;
        const characterRange = action.characterRange;
        const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
        return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      let factor = 1;
      for (const action of actions) {
        const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd));
        if (progress === -2) {
          factor = 1;
          continue;
        }
        if (progress === null || progress < 0 || action.fontSizeFrom === void 0 || action.fontSizeTo === void 0) continue;
        factor = ((_b = action.fontSizeFormulaSamples) == null ? void 0 : _b.length) ? animationKeyframeValue(action.fontSizeFormulaSamples, "lin", progress) : ((_c = action.fontSizeKeyframes) == null ? void 0 : _c.length) ? animationKeyframeValue(action.fontSizeKeyframes, action.fontSizeKeyframeMode, progress) : action.fontSizeFrom + (action.fontSizeTo - action.fontSizeFrom) * progress;
      }
      const baseFontWeight = run.bold ? 700 : 400;
      let weight;
      const weightActions = ((_d = animationsByTarget.value.get(element.id)) == null ? void 0 : _d.filter((action) => {
        if (action.effect !== "fontWeight") return false;
        const paragraphRange = action.paragraphRange;
        const characterRange = action.characterRange;
        const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
        return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      for (const action of weightActions) {
        const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd));
        if (progress === -2) {
          weight = baseFontWeight;
          continue;
        }
        if (progress === null || progress < 0) continue;
        if (action.fontWeightReset) {
          weight = baseFontWeight;
          continue;
        }
        if (action.fontWeightFrom === void 0 || action.fontWeightTo === void 0) continue;
        weight = ((_e = action.fontWeightFormulaSamples) == null ? void 0 : _e.length) ? animationKeyframeValue(action.fontWeightFormulaSamples, "lin", progress) : ((_f = action.fontWeightKeyframes) == null ? void 0 : _f.length) ? animationKeyframeValue(action.fontWeightKeyframes, action.fontWeightKeyframeMode, progress) : action.fontWeightFrom + (action.fontWeightTo - action.fontWeightFrom) * progress;
      }
      let italic;
      let fontFamily;
      let underline;
      let lineThrough;
      let verticalAlign;
      let verticalAlignChanged = false;
      let textShadow;
      let textShadowChanged = false;
      let textEmbossEnabled = false;
      let textEmbossChanged = false;
      const baseFontFamily = script === "eastAsia" ? run.fontFamilyEastAsia : script === "complexScript" ? run.fontFamilyComplexScript : run.fontFamily;
      let textOutline;
      let textOutlineChanged = false;
      const baseTextOutline = run.outlineColor !== void 0 && run.outlineWidth !== void 0 ? `${run.outlineWidth * scale.value}px ${run.outlineColor}` : "0px transparent";
      const textStyleActions = ((_g = animationsByTarget.value.get(element.id)) == null ? void 0 : _g.filter((action) => {
        if (action.effect !== "textStyle") return false;
        const paragraphRange = action.paragraphRange;
        const characterRange = action.characterRange;
        const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
        return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      for (const action of textStyleActions) {
        const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd));
        let value;
        if (progress === -2 && action.textStyleProperty === "fontWeight") {
          weight = baseFontWeight;
          continue;
        } else if (progress === -2 && action.textStyleProperty === "textTransform") {
          verticalAlign = run.baselineOffsetEm === void 0 ? "baseline" : `${run.baselineOffsetEm}em`;
          verticalAlignChanged = true;
          continue;
        } else if (progress === -2 && action.textStyleProperty === "outline") {
          textOutline = baseTextOutline;
          textOutlineChanged = true;
          continue;
        } else if (progress === -2 && action.textStyleProperty === "textShadow") {
          textShadow = runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd);
          textShadowChanged = true;
          continue;
        } else if (progress === -2 && action.textStyleProperty === "emboss") {
          textEmbossEnabled = false;
          textEmbossChanged = true;
          continue;
        } else if (progress === -2) {
          if (action.textStyleProperty === "fontFamily") value = baseFontFamily || "";
          else if (action.textStyleProperty === "fontStyle") value = Boolean(run.italic);
          else if (action.textStyleProperty === "underline") value = Boolean(run.underline);
          else if (action.textStyleProperty === "lineThrough") value = Boolean(run.lineThrough);
        } else if (progress === null || progress < 0) continue;
        else if ((_h = action.textStyleKeyframes) == null ? void 0 : _h.length) {
          value = action.textStyleKeyframes[0].value;
          for (const frame of action.textStyleKeyframes) {
            if (frame.offset > progress) break;
            value = frame.value;
          }
        } else if (progress >= 1) value = action.textStyleTo;
        else value = action.textStyleFrom;
        if (value === void 0) continue;
        if (action.textStyleProperty === "fontWeight") weight = value === "bold" ? 700 : value === "normal" ? 400 : baseFontWeight;
        else if (action.textStyleProperty === "fontStyle") italic = value;
        else if (action.textStyleProperty === "fontFamily") fontFamily = cssFontFamily(String(value));
        else if (action.textStyleProperty === "underline") underline = value;
        else if (action.textStyleProperty === "lineThrough") lineThrough = value;
        else if (action.textStyleProperty === "textTransform") {
          verticalAlign = value === "sub" || value === "super" ? value : "baseline";
          verticalAlignChanged = true;
        } else if (action.textStyleProperty === "outline") {
          textOutline = value === true ? baseTextOutline : "0px transparent";
          textOutlineChanged = true;
        } else if (action.textStyleProperty === "textShadow") {
          textShadow = value === "none" ? "none" : runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd);
          textShadowChanged = true;
        } else if (action.textStyleProperty === "emboss") {
          textEmbossEnabled = value === "emboss";
          textEmbossChanged = true;
        }
      }
      const embossOffset = Math.max(0.5, (run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value / 12);
      const textDecorationLine = [(underline != null ? underline : run.underline) ? "underline" : void 0, (lineThrough != null ? lineThrough : run.lineThrough) ? "line-through" : void 0].filter(Boolean).join(" ");
      const textDecorationStyles = separateDecorations ? {
        ...underline === void 0 ? {} : { "--pptx-underline-line": underline ? "underline" : "none" },
        ...lineThrough === void 0 ? {} : { textDecorationLine: lineThrough ? "line-through" : "none" }
      } : underline === void 0 && lineThrough === void 0 ? {} : { textDecorationLine: textDecorationLine || "none" };
      return {
        ...factor === 1 ? {} : { fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value * factor}px` },
        ...weight === void 0 ? {} : { fontWeight: weight },
        ...fontFamily === void 0 ? {} : { fontFamily },
        ...italic === void 0 ? {} : { fontStyle: italic ? "italic" : "normal" },
        ...textDecorationStyles,
        ...verticalAlignChanged ? { verticalAlign } : {},
        ...textOutlineChanged ? { WebkitTextStroke: textOutline } : {},
        ...textShadowChanged ? { textShadow } : {},
        ...textEmbossChanged ? { filter: textEmbossEnabled ? `drop-shadow(-${embossOffset}px -${embossOffset}px 0 rgba(255, 255, 255, 0.65)) drop-shadow(${embossOffset}px ${embossOffset}px 0 rgba(0, 0, 0, 0.65))` : "none" } : {}
      };
    }
    function textTransformAnimationStyle(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
        if (![
          "scale",
          "rotation",
          "motion",
          "slide"
        ].includes(action.effect) || !(action.paragraphRange || action.characterRange || action.iteration)) return false;
        return (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      if (!actions.length) return {};
      let scaling = [1, 1];
      let rotation = 0;
      const [motionX, motionY] = animationMotion(element, paragraphIndex, characterStart, characterEnd);
      const [slideX, slideY] = animationSlide(element, paragraphIndex, characterStart, characterEnd);
      for (const action of actions) {
        const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd));
        if (progress === -2) {
          if (action.effect === "scale") scaling = [1, 1];
          else if (action.effect === "rotation") rotation = 0;
          continue;
        }
        if (progress === null || progress < 0) continue;
        if (action.effect === "scale" && action.scaleFrom && action.scaleTo) scaling = animationScaleValue(action, progress, iterationIndexFor(action, characterStart, characterEnd));
        else if (action.effect === "rotation" && action.rotationFrom !== void 0 && action.rotationTo !== void 0) {
          const delta = action.rotationTo - action.rotationFrom;
          const repeatOffset = delta * (action.reversePlayback ? -1 : 1) * completedAnimationRepeats(action, iterationIndexFor(action, characterStart, characterEnd));
          const angle = action.rotationFrom + delta * progress + repeatOffset;
          rotation = action.rotationRelative || action.rotationAdditive ? rotation + angle : angle;
        }
      }
      return {
        display: "inline-block",
        transform: [
          motionX || motionY ? `translate(${motionX * scale.value}px, ${motionY * scale.value}px)` : "",
          slideX || slideY ? `translate(${slideX * scale.value}px, ${slideY * scale.value}px)` : "",
          scaling[0] !== 1 || scaling[1] !== 1 ? `scale(${scaling[0]}, ${scaling[1]})` : "",
          rotation ? `rotate(${rotation}deg)` : ""
        ].filter(Boolean).join(" "),
        transformOrigin: "center center"
      };
    }
    function animationMotion(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const textRange = paragraphIndex !== void 0 && characterStart !== void 0 && characterEnd !== void 0;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "motion" && (textRange ? Boolean(action.paragraphRange || action.characterRange || action.iteration) && (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start) : !action.paragraphRange && !action.characterRange && !action.iteration))) || [];
      let value = [0, 0];
      for (const action of actions) {
        const iterationIndex = textRange ? iterationIndexFor(action, characterStart, characterEnd) : void 0;
        const progress = animationProgress(action, iterationIndex);
        if (progress === -2) {
          if (action.motionCoordinateProperty === "ppt_x") value = [0, value[1]];
          else if (action.motionCoordinateProperty === "ppt_y") value = [value[0], 0];
          else value = [0, 0];
          continue;
        }
        const samples = action.motionPathSamples;
        if (progress !== null && progress >= 0 && (samples == null ? void 0 : samples.length)) {
          const samplePosition = Math.min(samples.length - 1, progress * (samples.length - 1));
          const from2 = samples[Math.floor(samplePosition)];
          const to2 = samples[Math.ceil(samplePosition)];
          const ratio2 = samplePosition - Math.floor(samplePosition);
          const offset2 = motionAccumulationOffset(action, iterationIndex);
          const motion = [from2.x + (to2.x - from2.x) * ratio2 + offset2[0], from2.y + (to2.y - from2.y) * ratio2 + offset2[1]];
          if (action.motionCoordinateProperty === "ppt_x") value = [motion[0], value[1]];
          else if (action.motionCoordinateProperty === "ppt_y") value = [value[0], motion[1]];
          else value = motion;
          continue;
        }
        const path = action.motionPath;
        if (progress === null || progress < 0 || !(path == null ? void 0 : path.length)) continue;
        const offset = motionAccumulationOffset(action, iterationIndex);
        if (path.length === 1) {
          const point = path[0];
          if (point) value = [point.x + offset[0], point.y + offset[1]];
          continue;
        }
        const distance = (action.motionPathLength || 0) * progress;
        let low = 0, high = path.length;
        while (low < high) {
          const middle = low + high >>> 1;
          const point = path[middle];
          if (point && point.distance <= distance) low = middle + 1;
          else high = middle;
        }
        const from = path[low - 1], to = path[low];
        if (!from) continue;
        if (!to) {
          value = [from.x + offset[0], from.y + offset[1]];
          continue;
        }
        const span = to.distance - from.distance;
        const ratio = span ? (distance - from.distance) / span : 0;
        value = [from.x + (to.x - from.x) * ratio + offset[0], from.y + (to.y - from.y) * ratio + offset[1]];
      }
      return value;
    }
    function animationSlide(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      let value = [0, 0];
      const textRange = paragraphIndex !== void 0 || characterStart !== void 0;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((item) => {
        if (item.effect !== "slide") return false;
        const iterationIndex = characterStart !== void 0 && characterEnd !== void 0 ? iterationIndexFor(item, characterStart, characterEnd) : void 0;
        if (!textRange) return !item.paragraphRange && !item.characterRange && !item.iteration;
        return (Boolean(item.paragraphRange || item.characterRange) || iterationIndex !== void 0 && iterationIndex >= 0) && (!item.paragraphRange || paragraphIndex !== void 0 && paragraphIndex >= item.paragraphRange.start && paragraphIndex <= item.paragraphRange.end) && (!item.characterRange || characterStart !== void 0 && characterEnd !== void 0 && characterStart < item.characterRange.end && characterEnd > item.characterRange.start) && (!item.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      })) || [];
      for (const action of actions) {
        const progress = animationProgress(action, characterStart !== void 0 && characterEnd !== void 0 ? iterationIndexFor(action, characterStart, characterEnd) : void 0);
        if (progress === -2) {
          value = [0, 0];
          continue;
        }
        if (progress === null) continue;
        const phase = progress < 0 ? 0 : progress;
        const distance = action.direction === "in" ? 1 - phase : phase;
        value = action.slideFrom === "left" ? [-props.width * distance, 0] : action.slideFrom === "right" ? [props.width * distance, 0] : action.slideFrom === "top" ? [0, -props.height * distance] : action.slideFrom === "bottom" ? [0, props.height * distance] : [0, 0];
      }
      return value;
    }
    function animationTargetsRange(action, paragraphIndex, characterStart, characterEnd) {
      if (characterStart !== void 0 && characterEnd !== void 0) {
        const overlaps = Boolean(action.characterRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
        const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : void 0;
        return overlaps || iterationIndex !== void 0 && iterationIndex >= 0;
      }
      if (paragraphIndex !== void 0) return Boolean(action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
      return !action.paragraphRange && !action.characterRange && !action.iteration;
    }
    function animationIterationIndex(action, paragraphIndex, characterStart, characterEnd) {
      if (characterStart !== void 0 && characterEnd !== void 0) return iterationIndexFor(action, characterStart, characterEnd);
      return paragraphIndex === void 0 ? void 0 : iterationIndexFor(action, paragraphIndex, paragraphIndex + 1);
    }
    function animationClipMaskStyle(element, clipPath, textRange) {
      if (!textRange) return { clipPath };
      const mask = clipPathMaskImage(clipPath, element.width * scale.value, element.height * scale.value);
      if (!mask) return { clipPath };
      return {
        maskImage: mask,
        WebkitMaskImage: mask,
        maskSize: "100% 100%",
        WebkitMaskSize: "100% 100%",
        maskRepeat: "no-repeat",
        WebkitMaskRepeat: "no-repeat"
      };
    }
    function animationWipe(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        if (action.effect !== "wipe") return false;
        if (characterStart !== void 0 && characterEnd !== void 0) {
          const overlaps = Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
          const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : void 0;
          return overlaps || iterationIndex !== void 0 && iterationIndex >= 0;
        }
        return paragraphIndex === void 0 ? !action.paragraphRange && !action.characterRange : Boolean(action.paragraphRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
      })) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.wipeDirection)) return {};
      let progress = first.direction === "in" ? 0 : 100;
      let direction = first.wipeDirection;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        direction = action.wipeDirection || direction;
        progress = action.direction === "in" ? phase * 100 : (1 - phase) * 100;
      }
      if (reset) return {};
      const mask = `linear-gradient(to ${direction === "up" ? "top" : direction === "down" ? "bottom" : direction}, #000 var(--wipe-progress), transparent var(--wipe-progress))`;
      return {
        "--wipe-progress": `${progress}%`,
        maskImage: mask,
        WebkitMaskImage: mask
      };
    }
    function animationBlinds(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        if (action.effect !== "blinds") return false;
        if (paragraphIndex !== void 0) return Boolean(action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
        if (characterStart !== void 0 && characterEnd !== void 0) return Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
        return !action.characterRange && !action.paragraphRange;
      })) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.blindsOrientation)) return {};
      let progress = first.direction === "in" ? 0 : 1;
      let orientation = first.blindsOrientation;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action);
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        orientation = action.blindsOrientation || orientation;
        progress = action.direction === "in" ? phase : 1 - phase;
      }
      if (reset) return {};
      const stripeSize = `${Math.max(0, Math.min(progress, 1) * 12.5)}%`;
      const mask = `repeating-linear-gradient(${orientation === "horizontal" ? "to bottom" : "to right"}, #000 0 ${stripeSize}, transparent ${stripeSize} 12.5%)`;
      return {
        "--blinds-stripe-size": stripeSize,
        maskImage: mask,
        WebkitMaskImage: mask
      };
    }
    function animationChecker(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        if (action.effect !== "checker") return false;
        if (paragraphIndex !== void 0) return Boolean(action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
        if (characterStart !== void 0 && characterEnd !== void 0) return Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
        return !action.paragraphRange && !action.characterRange;
      })) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.checkerOrientation)) return {};
      let progress = first.direction === "in" ? 0 : 1;
      let orientation = first.checkerOrientation;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action);
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        orientation = action.checkerOrientation || orientation;
        progress = action.direction === "in" ? phase : 1 - phase;
      }
      if (reset) return {};
      const width = element.width * scale.value;
      const height = element.height * scale.value;
      const amount = Math.max(0, Math.min(progress, 1));
      if (paragraphIndex === void 0 && characterStart === void 0) return { clipPath: gridTransitionClipPath(width, height, orientation, amount, true) };
      const mask = gridTransitionMaskImage(width, height, orientation, amount, true);
      return {
        maskImage: mask,
        WebkitMaskImage: mask,
        maskSize: "100% 100%",
        WebkitMaskSize: "100% 100%",
        maskRepeat: "no-repeat",
        WebkitMaskRepeat: "no-repeat"
      };
    }
    function animationShapeMask(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "shape" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd))) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.shapeFilter) || !first.shapeDirection) return {};
      let current = {
        shape: first.shapeFilter,
        direction: first.shapeDirection,
        progress: first.direction === "in" ? 0 : 1
      };
      let reset = false;
      for (const action of actions) {
        if (!action.shapeFilter || !action.shapeDirection) continue;
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        current = {
          shape: action.shapeFilter,
          direction: action.shapeDirection,
          progress: action.direction === "in" ? phase : 1 - phase
        };
      }
      if (reset || !current) return {};
      return animationClipMaskStyle(element, shapeEffectClipPath(current.shape, element.width * scale.value, element.height * scale.value, current.progress, current.direction), paragraphIndex !== void 0 || characterStart !== void 0);
    }
    function animationDissolve(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "dissolve" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd))) || [];
      const first = actions[0];
      if (!first) return {};
      let progress = first.direction === "in" ? 0 : 1;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        progress = action.direction === "in" ? phase : 1 - phase;
      }
      if (reset) return {};
      return animationClipMaskStyle(element, dissolveClipPath(element.width * scale.value, element.height * scale.value, progress, element.id), paragraphIndex !== void 0 || characterStart !== void 0);
    }
    function animationWheel(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "wheel" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd))) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.wheelSpokes)) return {};
      let progress = first.direction === "in" ? 0 : 1;
      let spokes = first.wheelSpokes;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        spokes = action.wheelSpokes || spokes;
        progress = action.direction === "in" ? phase : 1 - phase;
      }
      if (reset) return {};
      return animationClipMaskStyle(element, wheelClipPath(element.width * scale.value, element.height * scale.value, spokes, progress), paragraphIndex !== void 0 || characterStart !== void 0);
    }
    function animationRandomBars(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "randomBars" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd))) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.randomBarOrientation)) return {};
      let progress = first.direction === "in" ? 0 : 1;
      let orientation = first.randomBarOrientation;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        orientation = action.randomBarOrientation || orientation;
        progress = action.direction === "in" ? phase : 1 - phase;
      }
      if (reset) return {};
      return animationClipMaskStyle(element, randomBarClipPath(element.width * scale.value, element.height * scale.value, orientation, progress, element.id), paragraphIndex !== void 0 || characterStart !== void 0);
    }
    function animationStrips(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "strips" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd))) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.stripsDirection)) return {};
      let progress = first.direction === "in" ? 0 : 1;
      let direction = first.stripsDirection;
      let reset = false;
      for (const action of actions) {
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        direction = action.stripsDirection || direction;
        progress = action.direction === "in" ? phase : 1 - phase;
      }
      if (reset) return {};
      return animationClipMaskStyle(element, stripsClipPath(element.width * scale.value, element.height * scale.value, direction, progress), paragraphIndex !== void 0 || characterStart !== void 0);
    }
    function animationBarn(element, paragraphIndex, characterStart, characterEnd) {
      var _a;
      const actions = ((_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => action.effect === "barn" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd))) || [];
      const first = actions[0];
      if (!(first == null ? void 0 : first.barnOrientation) || !first.barnMotion) return {};
      let orientation = first.barnOrientation;
      let motion = first.barnMotion;
      let progress = 0;
      let reset = false;
      for (const action of actions) {
        if (!action.barnOrientation || !action.barnMotion) continue;
        const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
        if (phase === -2) {
          reset = true;
          continue;
        }
        if (phase === null || phase < 0) continue;
        reset = false;
        orientation = action.barnOrientation;
        motion = action.barnMotion;
        progress = phase;
      }
      if (reset) return {};
      return { clipPath: barnClipPath(orientation, motion, progress) };
    }
    function animationCss(actions, characterStart, characterEnd) {
      var _a;
      if (props.thumbnail && !props.playbackPreview || !(actions == null ? void 0 : actions.length)) return {};
      const visibilityActions = actions.filter((action) => [
        "appear",
        "fade",
        "wipe",
        "strips",
        "barn",
        "slide",
        "visibility"
      ].includes(action.effect));
      const first = visibilityActions[0];
      if (!first) return {};
      let opacity = first.effect === "visibility" || first.effect === "wipe" || first.effect === "strips" || first.effect === "barn" || first.effect === "slide" || first.direction === "out" ? 1 : 0;
      let visibility = first.effect === "visibility" ? first.visibilityFrom || "visible" : first.direction === "out" ? "visible" : "hidden";
      for (const action of visibilityActions) {
        const progress = animationProgress(action, characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd));
        if (progress === -2) {
          opacity = 1;
          visibility = "visible";
          continue;
        }
        if (progress === null || progress < 0) continue;
        if (action.effect === "appear") {
          const atEnd = progress >= 1;
          visibility = action.direction === "in" ? atEnd ? "visible" : "hidden" : atEnd ? "hidden" : "visible";
          opacity = visibility === "visible" ? 1 : 0;
        } else if (action.effect === "visibility") {
          let value = action.visibilityFrom || "visible";
          if ((_a = action.visibilityKeyframes) == null ? void 0 : _a.length) for (const frame of action.visibilityKeyframes) {
            if (frame.offset > progress) break;
            value = frame.value;
          }
          else if (progress >= 1) value = action.visibilityTo || value;
          visibility = value;
        } else if (action.effect === "wipe" || action.effect === "strips" || action.effect === "barn") {
          visibility = action.direction === "in" ? progress > 0 ? "visible" : "hidden" : progress >= 1 ? "hidden" : "visible";
          opacity = 1;
        } else if (action.effect === "slide") {
          visibility = action.direction === "in" || progress < 1 ? "visible" : "hidden";
          opacity = 1;
        } else {
          const easedProgress2 = action.effectProgress === void 0 ? 1 - (1 - progress) ** 2 : progress;
          opacity = action.direction === "in" ? easedProgress2 : 1 - easedProgress2;
          visibility = action.direction === "in" || progress < 1 ? "visible" : "hidden";
        }
      }
      return {
        opacity,
        visibility,
        pointerEvents: visibility === "visible" && opacity > 0 ? "auto" : "none"
      };
    }
    function animationStyle(element) {
      const actions = animationsByTarget.value.get(element.id);
      if (!(actions == null ? void 0 : actions.length)) return {};
      const animatedStyle = animationCss(actions.filter((action) => !action.paragraphRange && !action.characterRange && !action.iteration));
      const opacity = animationOpacity(element);
      return {
        ...animatedStyle,
        opacity: (typeof animatedStyle.opacity === "number" ? animatedStyle.opacity : 1) * opacity,
        pointerEvents: opacity === 0 ? "none" : animatedStyle.pointerEvents,
        ...animationWipe(element),
        ...animationBlinds(element),
        ...animationChecker(element),
        ...animationShapeMask(element),
        ...animationDissolve(element),
        ...animationWheel(element),
        ...animationRandomBars(element),
        ...animationStrips(element),
        ...animationBarn(element)
      };
    }
    function paragraphAnimationStyle(element, index) {
      var _a;
      const actions = (_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        const range = action.paragraphRange;
        return range && action.effect !== "barn" && !action.iteration && index >= range.start && index <= range.end;
      });
      const [slideX, slideY] = animationSlide(element, index);
      return {
        ...animationCss(actions),
        ...slideX || slideY ? { transform: `translate(${slideX * scale.value}px, ${slideY * scale.value}px)` } : {},
        ...animationWipe(element, index),
        ...animationBlinds(element, index),
        ...animationChecker(element, index),
        ...animationBarn(element, index),
        ...animationShapeMask(element, index),
        ...animationDissolve(element, index),
        ...animationWheel(element, index),
        ...animationRandomBars(element, index),
        ...animationStrips(element, index)
      };
    }
    function characterAnimationStyle(element, paragraphIndex, start, end) {
      var _a;
      const actions = (_a = animationsByTarget.value.get(element.id)) == null ? void 0 : _a.filter((action) => {
        var _a2;
        const range = action.characterRange;
        const overlaps = range && start < range.end && end > range.start;
        if (action.iteration) return Boolean(overlaps || !range) && ((_a2 = iterationIndexFor(action, start, end)) != null ? _a2 : -1) >= 0;
        return Boolean(overlaps);
      });
      const animatedStyle = animationCss(actions == null ? void 0 : actions.filter((action) => action.effect !== "barn"), start, end);
      const textOpacity = animationOpacity(element, "style.opacity", paragraphIndex, start, end);
      const opacity = typeof animatedStyle.opacity === "number" ? animatedStyle.opacity * textOpacity : void 0;
      return {
        ...animatedStyle,
        ...opacity === void 0 ? {} : { opacity },
        ...animationWipe(element, void 0, start, end),
        ...animationBlinds(element, void 0, start, end),
        ...animationChecker(element, void 0, start, end),
        ...animationBarn(element, void 0, start, end),
        ...animationShapeMask(element, void 0, start, end),
        ...animationDissolve(element, void 0, start, end),
        ...animationWheel(element, void 0, start, end),
        ...animationRandomBars(element, void 0, start, end),
        ...animationStrips(element, void 0, start, end)
      };
    }
    function hasTrigger(elementId) {
      var _a;
      return !props.thumbnail && Boolean((_a = props.slide.triggeredAnimations) == null ? void 0 : _a.some((sequence) => sequence.trigger.type === "shape" && sequence.trigger.id === elementId));
    }
    function triggerSequenceIsActive(index, sequence, now) {
      return sequence.steps.some((actions, stepIndex) => {
        const key = triggerKey(index, stepIndex);
        const startedAt = triggerStartedAt.get(key);
        const endedAt = triggerStepEndedAt.get(key);
        if (startedAt === void 0 || !Number.isFinite(startedAt)) return false;
        if (endedAt !== void 0 && now >= endedAt) return false;
        return actions.some((action) => {
          var _a;
          const actionStart = startedAt + action.delayMs;
          if (now < actionStart) return true;
          if (action.effect === "media" && action.mediaActionKey) {
            if (completedMediaActions.has(action.mediaActionKey)) return false;
            if (action.mediaWaitForEnd) {
              if (blockedMediaPlaybackIds.value.includes(action.targetId)) return false;
              const media = Array.from(((_a = stage.value) == null ? void 0 : _a.querySelectorAll("[data-media-id]")) || []).find((item) => item.dataset.mediaId === action.targetId);
              return !media || !media.paused;
            }
          }
          return now < actionStart + naturalAnimationDuration(action);
        });
      });
    }
    function startTriggeredAnimations(indices, startedAt = performance.now(), refresh = true) {
      if (props.triggerPlayback !== void 0) return;
      const sequences = props.slide.triggeredAnimations || [];
      if (!indices.length) return;
      for (const index of indices) {
        const sequence = sequences[index];
        if (!sequence) continue;
        const count = sequence.steps.length;
        if (!count) continue;
        let step = triggerSteps.get(index) || 0;
        if (step >= count) {
          if (sequence.restart === "never") continue;
          if (sequence.restart === "whenNotActive" && triggerSequenceIsActive(index, sequence, performance.now())) continue;
          for (let previous = 0; previous < count; previous++) {
            const key2 = triggerKey(index, previous);
            triggerStartedAt.delete(key2);
            triggerStepEndedAt.delete(key2);
            triggerPlaybackByKey.delete(key2);
          }
          timelineActions.value.forEach((action, actionIndex) => {
            if (action.triggerSequence === index) executedMediaActions.delete(actionIndex);
          });
          step = 0;
        }
        if (step > 0 && !sequence.concurrent) endTimelineStep(index, step - 1, startedAt);
        const key = triggerKey(index, step);
        const playback = {
          sequence: index,
          step,
          startedAt: performance.timeOrigin + startedAt
        };
        triggerStartedAt.set(key, startedAt);
        triggerPlaybackByKey.set(key, playback);
        triggerSteps.set(index, step + 1);
      }
      if (refresh) refreshTimeline();
      emit("trigger-playback", [...triggerPlaybackByKey.values()].sort((a, b) => a.sequence - b.sequence || a.step - b.step));
    }
    function runTimingNodeTriggeredAnimations(now) {
      var _a;
      if (props.thumbnail || props.triggerPlayback !== void 0) return Infinity;
      const due = [];
      let next = Infinity;
      (_a = props.slide.triggeredAnimations) == null ? void 0 : _a.forEach((sequence, index) => {
        var _a2;
        const trigger = sequence.trigger;
        if (trigger.type !== "timeNode" || firedTimingNodeTriggers.has(index)) return;
        const sourceStartedAt = trigger.sourceStep === void 0 ? automaticStartedAt : stepStartedAt.get(trigger.sourceStep);
        if (sourceStartedAt === void 0 || sourceStartedAt === Number.NEGATIVE_INFINITY) return;
        const startedAt = ((_a2 = trigger.event === "onEnd" && trigger.sourceStep !== void 0 ? mainStepEndedAt.get(trigger.sourceStep) : void 0) != null ? _a2 : sourceStartedAt + trigger.sourceDelayMs) + trigger.delayMs;
        if (!Number.isFinite(startedAt)) return;
        if (startedAt > now) next = Math.min(next, startedAt);
        else {
          firedTimingNodeTriggers.add(index);
          due.push({
            index,
            startedAt
          });
        }
      });
      for (const trigger of due) startTriggeredAnimations([trigger.index], trigger.startedAt, false);
      return next;
    }
    function runSlideTriggeredAnimations(triggerEvent) {
      if (props.thumbnail || props.triggerPlayback !== void 0) return;
      startTriggeredAnimations((props.slide.triggeredAnimations || []).flatMap((sequence, index) => sequence.trigger.type === "slide" && sequence.trigger.event === triggerEvent ? [index] : []));
    }
    function canActivateHyperlink(hyperlink) {
      return Boolean(hyperlink) && !props.thumbnail && !props.playbackPreview && props.triggerPlayback === void 0;
    }
    function hyperlinkHref(hyperlink) {
      return (hyperlink == null ? void 0 : hyperlink.kind) === "external" ? hyperlink.url : "#slide-link";
    }
    function usesRightToLeftColumns(element) {
      var _a;
      return element.textColumnsRightToLeft === true && ((_a = element.textColumnCount) != null ? _a : 1) > 1 && (!element.textOrientation || element.textOrientation === "horz");
    }
    function textFrameStyle(element) {
      var _a, _b;
      const verticalText = element.textOrientation === "vert" || element.textOrientation === "vert270";
      const eastAsianVerticalText = element.textOrientation === "eaVert";
      const wordArtVerticalText = element.textOrientation === "wordArtVert";
      const wordArtVerticalRtlText = element.textOrientation === "wordArtVertRtl";
      const mongolianVerticalText = element.textOrientation === "mongolianVert";
      const columnCount = (_a = element.textColumnCount) != null ? _a : 1;
      const verticalColumnWritingMode = columnCount > 1 && element.textOrientation && element.textOrientation !== "horz" && !wordArtVerticalRtlText ? element.textColumnsRightToLeft ? "vertical-rl" : "vertical-lr" : void 0;
      const rect = element.customTextRect;
      const fittedSize = rect ? autoFitDimensions.get(`${props.slide.id}:${element.id}`) : void 0;
      const ellipsisLineCount = supportsTextVerticalEllipsis(element) ? textEllipsisLineCounts.get(`${props.slide.id}:${element.id}`) : void 0;
      const margins = rect ? { ...element.margins } : element.margins;
      if (rect) {
        const availableWidth = Math.max(0, (rect.right - rect.left) * ((fittedSize == null ? void 0 : fittedSize.width) || element.width));
        const horizontalMargins = margins.left + margins.right;
        if (horizontalMargins > availableWidth) {
          const ratio = availableWidth / horizontalMargins;
          margins.left *= ratio;
          margins.right *= ratio;
        }
        const availableHeight = Math.max(0, (rect.bottom - rect.top) * ((fittedSize == null ? void 0 : fittedSize.height) || element.height));
        const verticalMargins = margins.top + margins.bottom;
        if (verticalMargins > availableHeight) {
          const ratio = availableHeight / verticalMargins;
          margins.top *= ratio;
          margins.bottom *= ratio;
        }
      }
      return {
        position: rect ? "absolute" : void 0,
        left: rect ? `${rect.left * 100}%` : void 0,
        top: rect ? `${rect.top * 100}%` : void 0,
        width: rect ? `${(rect.right - rect.left) * 100}%` : void 0,
        height: rect ? `${(rect.bottom - rect.top) * 100}%` : void 0,
        boxSizing: rect ? "border-box" : void 0,
        display: ellipsisLineCount ? "-webkit-box" : columnCount > 1 ? "block" : void 0,
        WebkitBoxOrient: ellipsisLineCount ? "vertical" : void 0,
        WebkitBoxPack: ellipsisLineCount ? element.verticalAlign === "middle" ? "center" : element.verticalAlign === "bottom" ? "end" : "start" : void 0,
        WebkitLineClamp: ellipsisLineCount,
        columnCount: columnCount > 1 ? columnCount : void 0,
        columnGap: columnCount > 1 ? `${((_b = element.textColumnSpacing) != null ? _b : 0) * scale.value}px` : void 0,
        columnFill: columnCount > 1 ? "auto" : void 0,
        padding: `${margins.top * scale.value}px ${margins.right * scale.value}px ${margins.bottom * scale.value}px ${margins.left * scale.value}px`,
        justifyContent: element.verticalAlign === "middle" ? "center" : element.verticalAlign === "bottom" ? "flex-end" : "flex-start",
        overflowX: element.textHorizontalOverflow === void 0 ? void 0 : element.textHorizontalOverflow === "clip" ? "clip" : "visible",
        overflowY: ellipsisLineCount ? "hidden" : element.textVerticalOverflow === void 0 ? void 0 : element.textVerticalOverflow === "overflow" ? "visible" : "clip",
        whiteSpace: element.textWrap === "none" ? "nowrap" : void 0,
        writingMode: verticalColumnWritingMode != null ? verticalColumnWritingMode : element.textOrientation === "vert270" || mongolianVerticalText ? "vertical-lr" : verticalText || eastAsianVerticalText || wordArtVerticalText || wordArtVerticalRtlText ? "vertical-rl" : void 0,
        textOrientation: verticalText ? "sideways" : eastAsianVerticalText || wordArtVerticalText || wordArtVerticalRtlText || mongolianVerticalText ? "upright" : void 0,
        direction: element.textOrientation === "vert270" || usesRightToLeftColumns(element) ? "rtl" : void 0
      };
    }
    function effectiveTextFontScale(element) {
      var _a, _b;
      return (_b = (_a = normalAutoFitScales.get(`${props.slide.id}:${element.id}`)) != null ? _a : element.textFontScale) != null ? _b : 1;
    }
    function tableCellAutoFitKey(elementId, rowIndex, cellIndex) {
      return `${props.slide.id}:${elementId}:${rowIndex}:${cellIndex}`;
    }
    function effectiveTableCellFontScale(element, rowIndex, cellIndex, cell) {
      var _a, _b;
      return (_b = (_a = normalAutoFitScales.get(tableCellAutoFitKey(element.id, rowIndex, cellIndex))) != null ? _a : cell.textFontScale) != null ? _b : 1;
    }
    function geometryStyle(element) {
      const geometry = element.geometry;
      if ([
        "ellipse",
        "flowChartConnector",
        "flowChartSummingJunction",
        "flowChartOr"
      ].includes(geometry)) return { borderRadius: "50%" };
      if (geometry === "roundRect" || geometry === "flowChartAlternateProcess") {
        const [x, y] = roundRectRadii(element);
        return { borderRadius: `${x / 10}% / ${y / 10}%` };
      }
      const clipPaths = {};
      return clipPaths[geometry] ? { clipPath: clipPaths[geometry] } : {};
    }
    function runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd) {
      var _a, _b;
      const shadows = ((_a = run.shadows) == null ? void 0 : _a.length) ? run.shadows : run.shadow ? [run.shadow] : [];
      const glows = ((_b = run.glows) == null ? void 0 : _b.length) ? run.glows : run.glow ? [run.glow] : [];
      const effects = shadows.map((shadow) => {
        const shadowColor = animationColor(element, "shadow.color", paragraphIndex, characterStart, characterEnd) || shadow.color;
        return shadowCss(shadow, animatedRunColor(shadowColor, animationOpacity(element, "shadow.opacity", paragraphIndex, characterStart, characterEnd)) || shadowColor);
      });
      effects.push(...glows.map((glow) => `0 0 ${glow.radius * scale.value}px ${glow.color}`));
      return effects.length ? effects.join(", ") : void 0;
    }
    function runStyle(element, run, script, fontScale, paragraphIndex, characterStart, characterEnd, separateDecorations = false) {
      var _a, _b;
      const requestedFont = script === "eastAsia" ? run.fontFamilyEastAsia : script === "complexScript" ? run.fontFamilyComplexScript : run.fontFamily;
      const fillOpacity = animationOpacity(element, "fill.opacity", paragraphIndex, characterStart, characterEnd);
      const strokeOpacity = animationOpacity(element, "stroke.opacity", paragraphIndex, characterStart, characterEnd);
      const textDecorationLine = separateDecorations ? run.lineThrough ? "line-through" : void 0 : [run.underline ? "underline" : void 0, run.lineThrough ? "line-through" : void 0].filter(Boolean).join(" ");
      const textDecorationStyle = separateDecorations ? run.lineThroughStyle : run.underline && run.lineThrough ? (_a = run.underlineStyle) != null ? _a : run.lineThroughStyle === "double" ? "solid" : run.lineThroughStyle : (_b = run.underlineStyle) != null ? _b : run.lineThroughStyle;
      return {
        fontFamily: cssFontFamily(requestedFont || run.fontFamily),
        fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value}px`,
        letterSpacing: run.letterSpacingPt === void 0 ? void 0 : `${run.letterSpacingPt * fontScale * EMU_PER_POINT * scale.value}px`,
        verticalAlign: run.baselineOffsetEm === void 0 ? void 0 : `${run.baselineOffsetEm}em`,
        color: animatedRunColor(run.color, fillOpacity),
        "--pptx-text-pattern": run.fillPattern ? applyColorOpacity(run.fillPattern, fillOpacity) : void 0,
        "--pptx-text-pattern-color": animatedRunColor(run.color, fillOpacity),
        WebkitTextStroke: run.outlineColor !== void 0 && run.outlineWidth !== void 0 ? `${run.outlineWidth * scale.value}px ${animatedRunColor(run.outlineColor, strokeOpacity)}` : void 0,
        fontWeight: run.bold ? 700 : void 0,
        fontStyle: run.italic ? "italic" : void 0,
        ...separateDecorations ? { "--pptx-underline-line": run.underline ? "underline" : "none" } : {},
        textDecorationLine: textDecorationLine || void 0,
        textDecorationStyle,
        textDecorationColor: !separateDecorations && run.underline ? run.underlineColor : void 0,
        textDecorationThickness: !separateDecorations && run.underline && run.underlineThicknessEmu ? `${run.underlineThicknessEmu * scale.value}px` : void 0,
        textShadow: runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd)
      };
    }
    function textUnderlineStyle(run) {
      return {
        textDecorationLine: "var(--pptx-underline-line, none)",
        textDecorationStyle: run.underlineStyle,
        textDecorationColor: run.underlineColor,
        textDecorationThickness: run.underlineThicknessEmu ? `${run.underlineThicknessEmu * scale.value}px` : void 0
      };
    }
    function runCharacterOffset(element, paragraphIndex, runIndex) {
      var _a;
      let offset = 0;
      for (let index = 0; index < paragraphIndex; index++) {
        const paragraph2 = element.paragraphs[index];
        if (paragraph2) offset += paragraph2.runs.reduce((sum, run) => sum + run.text.length, 0) + 1;
      }
      const paragraph = element.paragraphs[paragraphIndex];
      if (paragraph) for (let index = 0; index < runIndex; index++) offset += ((_a = paragraph.runs[index]) == null ? void 0 : _a.text.length) || 0;
      return offset;
    }
    function separateTextDecorations(element, run, paragraphIndex, characterStart, characterEnd) {
      if (run.underline && run.lineThrough) return true;
      return (animationsByTarget.value.get(element.id) || []).some((action) => {
        if (action.effect !== "textStyle" || action.textStyleProperty !== "underline" && action.textStyleProperty !== "lineThrough") return false;
        const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
        return (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
      });
    }
    function runSegments(run, element, paragraphIndex, runIndex) {
      var _a;
      const runStart = runCharacterOffset(element, paragraphIndex, runIndex);
      const characterRanges = (animationsByTarget.value.get(element.id) || []).flatMap((action) => {
        var _a2;
        return [...action.characterRange ? [action.characterRange] : [], ...((_a2 = action.iteration) == null ? void 0 : _a2.ranges) || []];
      });
      if (!characterRanges.length) {
        const segments2 = [];
        let segmentStart = 0;
        let segmentScript;
        let previousScript2 = "latin";
        let offset2 = 0;
        for (const character of run.text) {
          const script = characterScript(character, previousScript2);
          if (segmentScript && script !== segmentScript) {
            segments2.push({
              text: run.text.slice(segmentStart, offset2),
              script: segmentScript,
              start: runStart + segmentStart,
              end: runStart + offset2
            });
            segmentStart = offset2;
          }
          segmentScript = script;
          previousScript2 = script;
          offset2 += character.length;
        }
        if (segmentScript && offset2 > segmentStart) segments2.push({
          text: run.text.slice(segmentStart),
          script: segmentScript,
          start: runStart + segmentStart,
          end: runStart + offset2
        });
        return segments2.map((segment) => ({
          ...segment,
          separateDecorations: separateTextDecorations(element, run, paragraphIndex, segment.start, segment.end)
        }));
      }
      const characters = [];
      const cuts = /* @__PURE__ */ new Set([0, run.text.length]);
      let previousScript = "latin";
      let offset = 0;
      for (const character of run.text) {
        const script = characterScript(character, previousScript);
        const end = offset + character.length;
        if (characters.length && ((_a = characters.at(-1)) == null ? void 0 : _a.script) !== script) cuts.add(offset);
        characters.push({
          text: character,
          script,
          start: offset,
          end
        });
        previousScript = script;
        offset = end;
      }
      const boundaries = /* @__PURE__ */ new Set([0]);
      for (const character of characters) boundaries.add(character.end);
      for (const range of characterRanges) for (const boundary of [range.start, range.end]) {
        const local = boundary - runStart;
        if (local > 0 && local < run.text.length && boundaries.has(local)) cuts.add(local);
      }
      const sortedCuts = [...cuts].sort((left, right) => left - right);
      const segments = [];
      let characterIndex = 0;
      for (let index = 1; index < sortedCuts.length; index++) {
        const start = sortedCuts[index - 1];
        const end = sortedCuts[index];
        while (characters[characterIndex] && characters[characterIndex].end <= start) characterIndex++;
        const first = characters[characterIndex];
        if (!first || end <= start) continue;
        segments.push({
          text: run.text.slice(start, end),
          script: first.script,
          start: runStart + start,
          end: runStart + end
        });
      }
      return segments.map((segment) => ({
        ...segment,
        separateDecorations: separateTextDecorations(element, run, paragraphIndex, segment.start, segment.end)
      }));
    }
    function paragraphBulletStyle(paragraph) {
      var _a;
      const size = (_a = paragraph.bullet) == null ? void 0 : _a.size;
      if (!size) return {};
      return { fontSize: size.unit === "ratio" ? `${size.value * 100}%` : `${size.value * EMU_PER_POINT * scale.value}px` };
    }
    function cssSpacing(spacing) {
      if (!spacing) return void 0;
      return spacing.unit === "emu" ? `${spacing.value * scale.value}px` : `${spacing.value}em`;
    }
    function paragraphStyle(paragraph, fontScale = 1, lineSpaceReduction = 0, fallbackDirection) {
      var _a;
      const lineSpacing = paragraph.lineSpacing;
      return {
        fontSize: `${(paragraph.runs.reduce((size, run) => Math.max(size, run.fontSizePt || 18), 0) || 18) * fontScale * EMU_PER_POINT * scale.value}px`,
        textAlign: paragraph.align,
        direction: (_a = paragraph.direction) != null ? _a : fallbackDirection,
        marginLeft: paragraph.marginLeft === void 0 ? void 0 : `${paragraph.marginLeft * scale.value}px`,
        marginRight: paragraph.marginRight === void 0 ? void 0 : `${paragraph.marginRight * scale.value}px`,
        textIndent: paragraph.indent === void 0 ? void 0 : `${paragraph.indent * scale.value}px`,
        marginTop: cssSpacing(paragraph.spaceBefore),
        marginBottom: cssSpacing(paragraph.spaceAfter),
        lineHeight: (lineSpacing == null ? void 0 : lineSpacing.unit) === "ratio" ? lineSpacing.value * (1 - lineSpaceReduction) : cssSpacing(lineSpacing)
      };
    }
    function tableCellStyle(cell) {
      const border = (side) => {
        const value = cell.borders[side];
        return value ? `${Math.max(value.width * scale.value, 1)}px solid ${value.color}` : void 0;
      };
      return {
        background: cell.fill === "transparent" ? void 0 : cell.fill,
        padding: `${cell.margins.top * scale.value}px ${cell.margins.right * scale.value}px ${cell.margins.bottom * scale.value}px ${cell.margins.left * scale.value}px`,
        verticalAlign: cell.verticalAlign,
        borderTop: border("top"),
        borderRight: border("right"),
        borderBottom: border("bottom"),
        borderLeft: border("left")
      };
    }
    function tableCellTextStyle(cell, rowHeight) {
      var _a, _b;
      const columnCount = (_a = cell.textColumnCount) != null ? _a : 1;
      return columnCount > 1 || cell.textAutoFitDynamic ? {
        height: rowHeight > 0 ? `${Math.max(0, rowHeight - cell.margins.top - cell.margins.bottom) * scale.value}px` : "100%",
        boxSizing: "border-box",
        minHeight: 0,
        overflow: "hidden",
        ...columnCount > 1 ? {
          columnCount,
          columnGap: `${((_b = cell.textColumnSpacing) != null ? _b : 0) * scale.value}px`,
          columnFill: "auto",
          direction: cell.textColumnsRightToLeft ? "rtl" : void 0
        } : {}
      } : {};
    }
    return (_ctx, _push, _parent, _attrs) => {
      _push(`<div${ssrRenderAttrs(mergeProps({
        ref_key: "stage",
        ref: stage,
        class: ["slide-frame", { "is-thumbnail": __props.thumbnail }],
        style: frameStyle.value,
        role: hasInteractiveMedia.value ? "group" : "img",
        "aria-label": `Slide ${__props.number}`
      }, _attrs))} data-v-52fe89f0><!--[-->`);
      ssrRenderList(__props.slide.elements, (element) => {
        var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y, _z, _A, _B, _C, _D, _E, _F, _G, _H, _I, _J, _K, _L, _M, _N, _O, _P, _Q, _R, _S, _T, _U, _V, _W, _X, _Y, _Z, __, _$, _aa, _ba;
        _push(`<div${ssrRenderAttr("data-element-id", element.id)}${ssrRenderAttr("data-element-name", element.name)}${ssrRenderAttr("data-element-kind", element.kind)} class="${ssrRenderClass([{
          "slide-shape": element.kind === "shape",
          "shape-line": element.kind === "line",
          "slide-placeholder": element.kind === "placeholder",
          "is-animation-trigger": hasTrigger(element.id),
          "is-hyperlink": canActivateHyperlink(element.hyperlink)
        }, "slide-element"])}"${ssrRenderAttr("role", canActivateHyperlink(element.hyperlink) ? "link" : hasTrigger(element.id) ? "button" : void 0)}${ssrRenderAttr("tabindex", canActivateHyperlink(element.hyperlink) || hasTrigger(element.id) ? 0 : void 0)}${ssrRenderAttr("aria-label", ((_a = element.hyperlink) == null ? void 0 : _a.tooltip) || (hasTrigger(element.id) ? `Play animation from ${element.name}` : void 0))}${ssrRenderAttr("data-pptx-hyperlink", canActivateHyperlink(element.hyperlink) ? JSON.stringify(element.hyperlink) : void 0)} style="${ssrRenderStyle([
          elementStyle(element),
          geometryStyle(element),
          animationStyle(element)
        ])}"${ssrRenderAttr("title", ((_b = element.hyperlink) == null ? void 0 : _b.tooltip) || element.name)} data-v-52fe89f0>`);
        if (usesSvgOutline(element)) {
          _push(`<svg class="preset-outline" viewBox="0 0 1000 1000" preserveAspectRatio="none" style="${ssrRenderStyle(usesSvgPolygon(element) ? customGeometryStyle(element) : void 0)}" aria-hidden="true" data-v-52fe89f0><defs data-v-52fe89f0><!--[-->`);
          ssrRenderList(customGeometryGradients(element).filter((item) => usesSvgPolygon(element) || item.id === customGradientId(element, "stroke")), (gradient) => {
            _push(`<!--[-->`);
            if (gradient.kind === "linear") {
              _push(`<linearGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("x1", gradient.x1)}${ssrRenderAttr("y1", gradient.y1)}${ssrRenderAttr("x2", gradient.x2)}${ssrRenderAttr("y2", gradient.y2)} data-v-52fe89f0><!--[-->`);
              ssrRenderList(gradient.stops, (stop, index) => {
                _push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-52fe89f0></stop>`);
              });
              _push(`<!--]--></linearGradient>`);
            } else {
              _push(`<radialGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("cx", gradient.cx)}${ssrRenderAttr("cy", gradient.cy)}${ssrRenderAttr("r", gradient.r)}${ssrRenderAttr("fx", gradient.fx)}${ssrRenderAttr("fy", gradient.fy)}${ssrRenderAttr("fr", gradient.fr)}${ssrRenderAttr("gradientTransform", gradient.gradientTransform)}${ssrRenderAttr("color-interpolation", gradient.colorInterpolation)} data-v-52fe89f0><!--[-->`);
              ssrRenderList(gradient.stops, (stop, index) => {
                _push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-52fe89f0></stop>`);
              });
              _push(`<!--]--></radialGradient>`);
            }
            _push(`<!--]-->`);
          });
          _push(`<!--]--></defs><path${ssrRenderAttr("d", presetOutlinePath(element))}${ssrRenderAttr("fill", usesSvgPolygon(element) ? customGeometryFill(element, {
            fill: true}) : "none")}${ssrRenderAttr("stroke", presetOutlineDetailPaths(element).length ? "none" : customGeometryStroke(element, {
            stroke: true
          }))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}" vector-effect="non-scaling-stroke" data-v-52fe89f0></path><!--[-->`);
          ssrRenderList(presetOutlineDetailPaths(element), (detailPath, index) => {
            _push(`<path${ssrRenderAttr("d", detailPath)} fill="none"${ssrRenderAttr("stroke", customGeometryStroke(element, {
              stroke: true
            }))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}" vector-effect="non-scaling-stroke" data-v-52fe89f0></path>`);
          });
          _push(`<!--]--></svg>`);
        } else _push(`<!---->`);
        if (element.kind === "shape" && ((_c = element.customPaths) == null ? void 0 : _c.length)) {
          _push(`<svg class="custom-geometry" viewBox="0 0 1000 1000" preserveAspectRatio="none" style="${ssrRenderStyle(customGeometryStyle(element))}" aria-hidden="true" data-v-52fe89f0><defs data-v-52fe89f0><!--[-->`);
          ssrRenderList(customGeometryGradients(element), (gradient) => {
            _push(`<!--[-->`);
            if (gradient.kind === "linear") {
              _push(`<linearGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("x1", gradient.x1)}${ssrRenderAttr("y1", gradient.y1)}${ssrRenderAttr("x2", gradient.x2)}${ssrRenderAttr("y2", gradient.y2)} data-v-52fe89f0><!--[-->`);
              ssrRenderList(gradient.stops, (stop, stopIndex) => {
                _push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-52fe89f0></stop>`);
              });
              _push(`<!--]--></linearGradient>`);
            } else {
              _push(`<radialGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("cx", gradient.cx)}${ssrRenderAttr("cy", gradient.cy)}${ssrRenderAttr("r", gradient.r)}${ssrRenderAttr("fx", gradient.fx)}${ssrRenderAttr("fy", gradient.fy)}${ssrRenderAttr("fr", gradient.fr)}${ssrRenderAttr("gradientTransform", gradient.gradientTransform)}${ssrRenderAttr("color-interpolation", gradient.colorInterpolation)} data-v-52fe89f0><!--[-->`);
              ssrRenderList(gradient.stops, (stop, stopIndex) => {
                _push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-52fe89f0></stop>`);
              });
              _push(`<!--]--></radialGradient>`);
            }
            _push(`<!--]-->`);
          });
          _push(`<!--]--><!--[-->`);
          ssrRenderList(element.customPaths, (path, pathIndex) => {
            _push(`<!--[-->`);
            if (customGeometryFillFocus(element, path)) _push(`<clipPath${ssrRenderAttr("id", customGradientFocusClipId(element, pathIndex))} clipPathUnits="userSpaceOnUse" data-v-52fe89f0><path${ssrRenderAttr("d", path.d)} data-v-52fe89f0></path></clipPath>`);
            else _push(`<!---->`);
            _push(`<!--]-->`);
          });
          _push(`<!--]--></defs><!--[-->`);
          ssrRenderList(element.customPaths, (path, pathIndex) => {
            _push(`<!--[--><path${ssrRenderAttr("d", path.d)}${ssrRenderAttr("fill", customGeometryFill(element, path))}${ssrRenderAttr("stroke", customGeometryStroke(element, path))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}" vector-effect="non-scaling-stroke" data-v-52fe89f0></path><!--[-->`);
            ssrRenderList([customGeometryFillFocus(element, path)], (focus) => {
              _push(`<!--[-->`);
              if ((focus == null ? void 0 : focus.kind) === "ellipse") _push(`<ellipse${ssrRenderAttr("cx", focus.cx)}${ssrRenderAttr("cy", focus.cy)}${ssrRenderAttr("rx", focus.rx)}${ssrRenderAttr("ry", focus.ry)}${ssrRenderAttr("fill", focus.color)}${ssrRenderAttr("clip-path", `url(#${customGradientFocusClipId(element, pathIndex)})`)} data-v-52fe89f0></ellipse>`);
              else if (focus) _push(`<rect${ssrRenderAttr("x", focus.x)}${ssrRenderAttr("y", focus.y)}${ssrRenderAttr("width", focus.width)}${ssrRenderAttr("height", focus.height)}${ssrRenderAttr("fill", focus.color)}${ssrRenderAttr("clip-path", `url(#${customGradientFocusClipId(element, pathIndex)})`)} data-v-52fe89f0></rect>`);
              else _push(`<!---->`);
              _push(`<!--]-->`);
            });
            _push(`<!--]--><!--]-->`);
          });
          _push(`<!--]--></svg>`);
        } else _push(`<!---->`);
        if (element.kind === "video" && element.mediaUrl && !__props.thumbnail) {
          _push(`<div class="image-viewport" data-v-52fe89f0><video class="slide-video"${ssrRenderAttr("data-media-id", element.id)}${ssrRenderAttr("data-media-slide-count", mediaSlideCounts.value.get(element.id) || void 0)}${ssrRenderAttr("src", element.mediaUrl)}${ssrRenderAttr("poster", element.imageUrl)} controls playsinline preload="metadata"${ssrRenderAttr("aria-label", element.name)} style="${ssrRenderStyle(imageStyle(element))}" data-v-52fe89f0></video>`);
          if (invalidMediaTrimIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>Trim range exceeds media duration; full media is used.</span>`);
          else _push(`<!---->`);
          if (invalidMediaFadeIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>Media fade duration exceeds the playable range; fades are ignored.</span>`);
          else _push(`<!---->`);
          if (blockedMediaPlaybackIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>The browser blocked automatic playback; use the media controls.</span>`);
          else _push(`<!---->`);
          _push(`</div>`);
        } else if (element.kind === "audio" && element.mediaUrl && !__props.thumbnail) {
          _push(`<div class="audio-viewport" data-v-52fe89f0><audio class="slide-audio"${ssrRenderAttr("data-media-id", element.id)}${ssrRenderAttr("data-media-slide-count", mediaSlideCounts.value.get(element.id) || void 0)}${ssrRenderAttr("src", element.mediaUrl)} controls preload="metadata"${ssrRenderAttr("aria-label", element.name)} data-v-52fe89f0></audio>`);
          if (invalidMediaTrimIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>Trim range exceeds media duration; full media is used.</span>`);
          else _push(`<!---->`);
          if (invalidMediaFadeIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>Media fade duration exceeds the playable range; fades are ignored.</span>`);
          else _push(`<!---->`);
          if (blockedMediaPlaybackIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>The browser blocked automatic playback; use the media controls.</span>`);
          else _push(`<!---->`);
          _push(`</div>`);
        } else if (element.kind === "chart" && element.chart) {
          _push(`<svg class="slide-chart"${ssrRenderAttr("data-chart-type", element.chart.type)}${ssrRenderAttr("data-radar-style", element.chart.radarStyle)} viewBox="0 0 1000 600" preserveAspectRatio="none" role="img"${ssrRenderAttr("aria-label", element.chart.title || element.name)} data-v-52fe89f0><title data-v-52fe89f0>${ssrInterpolate(element.chart.title || element.name)}</title><defs data-v-52fe89f0><clipPath${ssrRenderAttr("id", `chart-plot-${element.id}`)} data-v-52fe89f0><polygon${ssrRenderAttr("points", `${chartPlot(element.id).left},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top + chartPlot(element.id).height} ${chartPlot(element.id).left},${chartPlot(element.id).top + chartPlot(element.id).height}`)} data-v-52fe89f0></polygon></clipPath>`);
          if ((_d = chartVisuals.value.get(element.id)) == null ? void 0 : _d.legendClip) _push(`<clipPath${ssrRenderAttr("id", (_f = (_e = chartVisuals.value.get(element.id)) == null ? void 0 : _e.legendClip) == null ? void 0 : _f.id)} data-v-52fe89f0><rect${ssrRenderAttr("x", (_h = (_g = chartVisuals.value.get(element.id)) == null ? void 0 : _g.legendClip) == null ? void 0 : _h.x)}${ssrRenderAttr("y", (_j = (_i = chartVisuals.value.get(element.id)) == null ? void 0 : _i.legendClip) == null ? void 0 : _j.y)}${ssrRenderAttr("width", (_l = (_k = chartVisuals.value.get(element.id)) == null ? void 0 : _k.legendClip) == null ? void 0 : _l.width)}${ssrRenderAttr("height", (_n = (_m = chartVisuals.value.get(element.id)) == null ? void 0 : _m.legendClip) == null ? void 0 : _n.height)} data-v-52fe89f0></rect></clipPath>`);
          else _push(`<!---->`);
          _push(`</defs>`);
          if (element.chart.title) _push(`<text x="500" y="42" text-anchor="middle" class="chart-title" data-v-52fe89f0>${ssrInterpolate(element.chart.title)}</text>`);
          else _push(`<!---->`);
          if (element.chart.type === "pie" || element.chart.type === "doughnut") {
            _push(`<g data-v-52fe89f0><!--[-->`);
            ssrRenderList((_o = chartVisuals.value.get(element.id)) == null ? void 0 : _o.slices, (slice, index) => {
              _push(`<path${ssrRenderAttr("d", slice.d)}${ssrRenderAttr("fill", slice.color)}${ssrRenderAttr("fill-rule", element.chart.type === "doughnut" ? "evenodd" : "nonzero")} stroke="white" stroke-width="2" data-v-52fe89f0></path>`);
            });
            _push(`<!--]--></g>`);
          } else {
            _push(`<g data-v-52fe89f0>`);
            if (element.chart.type === "radar") {
              _push(`<!--[--><!--[-->`);
              ssrRenderList((_p = chartVisuals.value.get(element.id)) == null ? void 0 : _p.radarGrid, (grid, index) => {
                _push(`<path class="chart-major-gridline chart-radar-gridline"${ssrRenderAttr("d", grid.d)} fill="none" stroke="#d1d5db" stroke-width="1" data-v-52fe89f0></path>`);
              });
              _push(`<!--]--><!--[-->`);
              ssrRenderList((_q = chartVisuals.value.get(element.id)) == null ? void 0 : _q.radarSpokes, (spoke, index) => {
                _push(`<line${ssrRenderAttrs(mergeProps({
                  key: `radar-spoke-${index}`,
                  class: "chart-major-gridline chart-radar-spoke"
                }, { ref_for: true }, spoke, {
                  stroke: "#d1d5db",
                  "stroke-width": "1"
                }))} data-v-52fe89f0></line>`);
              });
              _push(`<!--]--><!--]-->`);
            } else {
              _push(`<!--[-->`);
              if (!element.chart.categoryAxisDeleted && element.chart.categoryAxisMajorGridlines) {
                _push(`<!--[-->`);
                ssrRenderList((_r = chartVisuals.value.get(element.id)) == null ? void 0 : _r.categoryGridlines, (grid, index) => {
                  _push(`<line${ssrRenderAttrs(mergeProps({
                    key: `category-grid-${index}`,
                    class: "chart-major-gridline chart-category-major-gridline",
                    "data-axis": "category"
                  }, { ref_for: true }, grid, {
                    stroke: "#d1d5db",
                    "stroke-width": "1"
                  }))} data-v-52fe89f0></line>`);
                });
                _push(`<!--]-->`);
              } else _push(`<!---->`);
              if (!((_s = element.chart.valueAxis) == null ? void 0 : _s.deleted) && ((_t = element.chart.valueAxis) == null ? void 0 : _t.majorGridlines)) {
                _push(`<!--[-->`);
                ssrRenderList((_u = chartVisuals.value.get(element.id)) == null ? void 0 : _u.ticks, (tick, index) => {
                  _push(`<line class="chart-major-gridline"${ssrRenderAttr("x1", element.chart.direction === "horizontal" ? tick.x : chartPlot(element.id).left)}${ssrRenderAttr("x2", element.chart.direction === "horizontal" ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width)}${ssrRenderAttr("y1", element.chart.direction === "horizontal" ? chartPlot(element.id).top : tick.y)}${ssrRenderAttr("y2", element.chart.direction === "horizontal" ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y)} stroke="#d1d5db" stroke-width="1" data-v-52fe89f0></line>`);
                });
                _push(`<!--]-->`);
              } else _push(`<!---->`);
              _push(`<!--[-->`);
              ssrRenderList((_v = chartVisuals.value.get(element.id)) == null ? void 0 : _v.secondaryValueAxes, (axis) => {
                _push(`<!--[-->`);
                if (!axis.deleted && axis.majorGridlines) {
                  _push(`<!--[-->`);
                  ssrRenderList(axis.ticks, (tick, index) => {
                    _push(`<line class="chart-major-gridline chart-secondary-gridline"${ssrRenderAttr("x1", axis.horizontal ? tick.x : chartPlot(element.id).left)}${ssrRenderAttr("x2", axis.horizontal ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width)}${ssrRenderAttr("y1", axis.horizontal ? chartPlot(element.id).top : tick.y)}${ssrRenderAttr("y2", axis.horizontal ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y)} stroke="#d1d5db" stroke-width="1" data-v-52fe89f0></line>`);
                  });
                  _push(`<!--]-->`);
                } else _push(`<!---->`);
                _push(`<!--]-->`);
              });
              _push(`<!--]-->`);
              if (["scatter", "bubble"].includes(element.chart.type) && !((_w = element.chart.xAxis) == null ? void 0 : _w.deleted) && ((_x = element.chart.xAxis) == null ? void 0 : _x.majorGridlines)) {
                _push(`<!--[-->`);
                ssrRenderList((_y = chartVisuals.value.get(element.id)) == null ? void 0 : _y.xTicks, (tick, index) => {
                  _push(`<line class="chart-major-gridline chart-x-gridline"${ssrRenderAttr("x1", tick.x)}${ssrRenderAttr("x2", tick.x)}${ssrRenderAttr("y1", chartPlot(element.id).top)}${ssrRenderAttr("y2", chartPlot(element.id).top + chartPlot(element.id).height)} stroke="#d1d5db" stroke-width="1" data-v-52fe89f0></line>`);
                });
                _push(`<!--]-->`);
              } else _push(`<!---->`);
              _push(`<!--]-->`);
            }
            _push(`<!--[-->`);
            ssrRenderList((_z = chartVisuals.value.get(element.id)) == null ? void 0 : _z.ticks, (tick, index) => {
              _push(`<!--[-->`);
              if (tick.showLabel) _push(`<text${ssrRenderAttr("x", tick.labelX)}${ssrRenderAttr("y", tick.labelY)}${ssrRenderAttr("text-anchor", tick.labelAnchor)} class="chart-label chart-value-axis-label" data-v-52fe89f0>${ssrInterpolate(tick.label)}</text>`);
              else _push(`<!---->`);
              _push(`<!--]-->`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_A = chartVisuals.value.get(element.id)) == null ? void 0 : _A.secondaryValueAxes, (axis) => {
              _push(`<!--[--><!--[-->`);
              ssrRenderList(axis.ticks, (tick, index) => {
                _push(`<!--[-->`);
                if (tick.showLabel) _push(`<text${ssrRenderAttr("data-axis-id", axis.id)}${ssrRenderAttr("x", tick.labelX)}${ssrRenderAttr("y", tick.labelY)}${ssrRenderAttr("text-anchor", tick.labelAnchor)} class="chart-label chart-value-axis-label chart-secondary-value-axis-label" data-v-52fe89f0>${ssrInterpolate(tick.label)}</text>`);
                else _push(`<!---->`);
                _push(`<!--]-->`);
              });
              _push(`<!--]--><!--]-->`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_B = chartVisuals.value.get(element.id)) == null ? void 0 : _B.xTicks, (tick, index) => {
              _push(`<!--[-->`);
              if (tick.showLabel) _push(`<text${ssrRenderAttr("x", tick.labelX)}${ssrRenderAttr("y", tick.labelY)}${ssrRenderAttr("text-anchor", tick.labelAnchor)} class="chart-label chart-x-value-axis-label" data-v-52fe89f0>${ssrInterpolate(tick.label)}</text>`);
              else _push(`<!---->`);
              _push(`<!--]-->`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_C = chartVisuals.value.get(element.id)) == null ? void 0 : _C.axisTicks, (tick, index) => {
              _push(`<line class="${ssrRenderClass(tick.minor ? "chart-axis-minor-tick" : "chart-axis-tick")}"${ssrRenderAttr("data-axis", tick.axis)}${ssrRenderAttr("x1", tick.x1)}${ssrRenderAttr("y1", tick.y1)}${ssrRenderAttr("x2", tick.x2)}${ssrRenderAttr("y2", tick.y2)} stroke="#64748b"${ssrRenderAttr("stroke-width", tick.minor ? 1 : 1.5)} data-v-52fe89f0></line>`);
            });
            _push(`<!--]--><g${ssrRenderAttr("clip-path", `url(#chart-plot-${element.id})`)} data-v-52fe89f0>`);
            if (element.chart.type !== "radar" && (["scatter", "bubble"].includes(element.chart.type) ? !((_D = element.chart.xAxis) == null ? void 0 : _D.deleted) : !element.chart.categoryAxisDeleted) && ((_E = chartVisuals.value.get(element.id)) == null ? void 0 : _E.baseline)) _push(`<line class="${ssrRenderClass([{ "chart-x-axis-line": ["scatter", "bubble"].includes(element.chart.type) }, "chart-category-axis-line"])}"${ssrRenderAttr("data-axis", ["scatter", "bubble"].includes(element.chart.type) ? "xValue" : "category")}${ssrRenderAttr("x1", (_F = chartVisuals.value.get(element.id)) == null ? void 0 : _F.baseline.x1)}${ssrRenderAttr("y1", (_G = chartVisuals.value.get(element.id)) == null ? void 0 : _G.baseline.y1)}${ssrRenderAttr("x2", (_H = chartVisuals.value.get(element.id)) == null ? void 0 : _H.baseline.x2)}${ssrRenderAttr("y2", (_I = chartVisuals.value.get(element.id)) == null ? void 0 : _I.baseline.y2)} stroke="#64748b" stroke-width="2" data-v-52fe89f0></line>`);
            else _push(`<!---->`);
            if (!((_J = element.chart.valueAxis) == null ? void 0 : _J.deleted) && ((_K = chartVisuals.value.get(element.id)) == null ? void 0 : _K.valueAxisLine)) _push(`<line class="chart-value-axis-line"${ssrRenderAttr("x1", (_M = (_L = chartVisuals.value.get(element.id)) == null ? void 0 : _L.valueAxisLine) == null ? void 0 : _M.x1)}${ssrRenderAttr("y1", (_O = (_N = chartVisuals.value.get(element.id)) == null ? void 0 : _N.valueAxisLine) == null ? void 0 : _O.y1)}${ssrRenderAttr("x2", (_Q = (_P = chartVisuals.value.get(element.id)) == null ? void 0 : _P.valueAxisLine) == null ? void 0 : _Q.x2)}${ssrRenderAttr("y2", (_S = (_R = chartVisuals.value.get(element.id)) == null ? void 0 : _R.valueAxisLine) == null ? void 0 : _S.y2)} stroke="#64748b" stroke-width="2" data-v-52fe89f0></line>`);
            else _push(`<!---->`);
            _push(`<!--[-->`);
            ssrRenderList((_T = chartVisuals.value.get(element.id)) == null ? void 0 : _T.secondaryValueAxes, (axis) => {
              _push(`<!--[-->`);
              if (!axis.deleted) _push(`<line class="chart-value-axis-line chart-secondary-value-axis-line"${ssrRenderAttr("data-axis-id", axis.id)}${ssrRenderAttr("x1", axis.line.x1)}${ssrRenderAttr("y1", axis.line.y1)}${ssrRenderAttr("x2", axis.line.x2)}${ssrRenderAttr("y2", axis.line.y2)} stroke="#64748b" stroke-width="2" data-v-52fe89f0></line>`);
              else _push(`<!---->`);
              _push(`<!--]-->`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_U = chartVisuals.value.get(element.id)) == null ? void 0 : _U.areas, (area, index) => {
              var _a2;
              _push(`<path class="chart-area"${ssrRenderAttr("d", area.d)}${ssrRenderAttr("fill", area.color)}${ssrRenderAttr("fill-opacity", (_a2 = area.opacity) != null ? _a2 : 1)}${ssrRenderAttr("stroke", area.color)} stroke-width="1" data-v-52fe89f0></path>`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_V = chartVisuals.value.get(element.id)) == null ? void 0 : _V.bars, (bar, index) => {
              _push(`<rect class="chart-bar"${ssrRenderAttr("x", bar.x)}${ssrRenderAttr("y", bar.y)}${ssrRenderAttr("width", bar.width)}${ssrRenderAttr("height", bar.height)}${ssrRenderAttr("fill", bar.color)}${ssrRenderAttr("stroke", bar.stroke)} stroke-width="1" data-v-52fe89f0></rect>`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_W = chartVisuals.value.get(element.id)) == null ? void 0 : _W.lines, (line, index) => {
              var _a2;
              _push(`<path class="chart-series-line"${ssrRenderAttr("d", line.d)} fill="none"${ssrRenderAttr("stroke", line.color)}${ssrRenderAttr("stroke-width", (_a2 = line.width) != null ? _a2 : 4)} stroke-linecap="round" data-v-52fe89f0></path>`);
            });
            _push(`<!--]--><!--[-->`);
            ssrRenderList((_X = chartVisuals.value.get(element.id)) == null ? void 0 : _X.markers, (marker, index) => {
              var _a2, _b2, _c2, _d2, _e2;
              _push(`<!--[-->`);
              if (!chartMarkerPath(marker)) _push(`<circle class="${ssrRenderClass([{ "chart-bubble-marker": element.chart.type === "bubble" }, "chart-line-marker"])}"${ssrRenderAttr("data-marker-shape", (_a2 = marker.symbol) != null ? _a2 : "circle")}${ssrRenderAttr("data-x", marker.x)}${ssrRenderAttr("data-y", marker.y)}${ssrRenderAttr("data-radius", (_b2 = marker.radius) != null ? _b2 : 5)}${ssrRenderAttr("cx", marker.x)}${ssrRenderAttr("cy", marker.y)}${ssrRenderAttr("r", marker.symbol === "dot" ? ((_c2 = marker.radius) != null ? _c2 : 5) * 0.45 : (_d2 = marker.radius) != null ? _d2 : 5)}${ssrRenderAttr("fill", marker.color)} data-v-52fe89f0></circle>`);
              else _push(`<path class="chart-line-marker"${ssrRenderAttr("data-marker-shape", marker.symbol)}${ssrRenderAttr("data-x", marker.x)}${ssrRenderAttr("data-y", marker.y)}${ssrRenderAttr("data-radius", (_e2 = marker.radius) != null ? _e2 : 5)}${ssrRenderAttr("d", chartMarkerPath(marker))}${ssrRenderAttr("fill", marker.color)} data-v-52fe89f0></path>`);
              _push(`<!--]-->`);
            });
            _push(`<!--]--></g><!--[-->`);
            ssrRenderList((_Y = chartVisuals.value.get(element.id)) == null ? void 0 : _Y.categories, (label, index) => {
              _push(`<text${ssrRenderAttr("x", label.x)}${ssrRenderAttr("y", label.y)}${ssrRenderAttr("text-anchor", label.anchor)} class="chart-label chart-category-label" data-v-52fe89f0>${ssrInterpolate(label.text)}</text>`);
            });
            _push(`<!--]--></g>`);
          }
          _push(`<!--[-->`);
          ssrRenderList((_Z = chartVisuals.value.get(element.id)) == null ? void 0 : _Z.dataLabels, (label, index) => {
            var _a2, _b2, _c2, _d2, _e2, _f2, _g2, _h2, _i2, _j2, _k2, _l2, _m2, _n2, _o2, _p2, _q2, _r2, _s2, _t2;
            _push(`<g data-v-52fe89f0>`);
            if (label.legendKey) {
              _push(`<g${ssrRenderAttr("data-key-id", chartDataLabelKeyId(element.id, index))}${ssrRenderAttr("fill", label.legendKey.color)} class="chart-data-label-key" data-v-52fe89f0>`);
              if (!label.legendKey.markerChart) _push(`<rect${ssrRenderAttr("x", (_a2 = chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index))) != null ? _a2 : label.x - 18)}${ssrRenderAttr("y", label.y - 12)} width="12" height="12" class="chart-data-label-key-swatch" data-v-52fe89f0></rect>`);
              else _push(`<!---->`);
              if (label.legendKey.line) _push(`<path${ssrRenderAttr("d", `M ${(_b2 = chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index))) != null ? _b2 : label.x - 18} ${label.y - 6} H ${((_c2 = chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index))) != null ? _c2 : label.x - 18) + 12}`)} fill="none"${ssrRenderAttr("stroke", label.legendKey.color)} stroke-width="2" stroke-linecap="round" class="chart-data-label-key-line" data-v-52fe89f0></path>`);
              else _push(`<!---->`);
              if (label.legendKey.symbol) {
                _push(`<!--[-->`);
                if (!chartMarkerPath({
                  x: ((_d2 = chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index))) != null ? _d2 : label.x - 18) + 6,
                  y: label.y - 6,
                  color: label.legendKey.color,
                  radius: 4,
                  symbol: label.legendKey.symbol
                })) _push(`<circle${ssrRenderAttr("cx", ((_e2 = chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index))) != null ? _e2 : label.x - 18) + 6)}${ssrRenderAttr("cy", label.y - 6)}${ssrRenderAttr("r", label.legendKey.symbol === "dot" ? 1.8 : 4)} class="chart-data-label-key-marker"${ssrRenderAttr("data-marker-shape", label.legendKey.symbol)} data-v-52fe89f0></circle>`);
                else _push(`<path${ssrRenderAttr("d", chartMarkerPath({
                  x: ((_f2 = chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index))) != null ? _f2 : label.x - 18) + 6,
                  y: label.y - 6,
                  color: label.legendKey.color,
                  radius: 4,
                  symbol: label.legendKey.symbol
                }))} class="chart-data-label-key-marker"${ssrRenderAttr("data-marker-shape", label.legendKey.symbol)} data-v-52fe89f0></path>`);
                _push(`<!--]-->`);
              } else _push(`<!---->`);
              _push(`</g>`);
            } else _push(`<!---->`);
            _push(`<text${ssrRenderAttr("x", label.x)}${ssrRenderAttr("y", label.y)}${ssrRenderAttr("text-anchor", label.anchor)} style="${ssrRenderStyle({
              fill: (_g2 = label.style) == null ? void 0 : _g2.color,
              fontFamily: ((_h2 = label.style) == null ? void 0 : _h2.fontFamily) ? unref(cssFontFamily)(label.style.fontFamily) : void 0,
              fontSize: ((_i2 = label.style) == null ? void 0 : _i2.fontSizePt) === void 0 ? void 0 : `${label.style.fontSizePt * 4 / 3}px`,
              fontWeight: ((_j2 = label.style) == null ? void 0 : _j2.bold) === void 0 ? void 0 : label.style.bold ? 700 : 400,
              fontStyle: ((_k2 = label.style) == null ? void 0 : _k2.italic) === void 0 ? void 0 : label.style.italic ? "italic" : "normal",
              letterSpacing: ((_l2 = label.style) == null ? void 0 : _l2.letterSpacingPt) === void 0 ? void 0 : `${label.style.letterSpacingPt * 4 / 3}px`,
              textDecorationLine: ((_m2 = label.style) == null ? void 0 : _m2.underline) === void 0 && ((_n2 = label.style) == null ? void 0 : _n2.lineThrough) === void 0 ? void 0 : [((_o2 = label.style) == null ? void 0 : _o2.underline) ? "underline" : void 0, ((_p2 = label.style) == null ? void 0 : _p2.lineThrough) ? "line-through" : void 0].filter(Boolean).join(" ") || "none",
              textDecorationStyle: ((_q2 = label.style) == null ? void 0 : _q2.underline) ? label.style.underlineStyle : ((_r2 = label.style) == null ? void 0 : _r2.lineThrough) ? label.style.lineThroughStyle : void 0,
              textDecorationColor: ((_s2 = label.style) == null ? void 0 : _s2.underline) ? label.style.underlineColor : void 0,
              textDecorationThickness: ((_t2 = label.style) == null ? void 0 : _t2.underline) && label.style.underlineThicknessEmu !== void 0 ? `${label.style.underlineThicknessEmu * 4 / (3 * EMU_PER_POINT)}px` : void 0
            })}" class="chart-data-label" data-v-52fe89f0><!--[-->`);
            ssrRenderList(label.lines, (line, lineIndex) => {
              var _a3;
              _push(`<tspan${ssrRenderAttr("x", label.x)}${ssrRenderAttr("y", label.y + (lineIndex - (label.lines.length - 1) / 2) * 22)}${ssrRenderAttr("dy", ((_a3 = label.style) == null ? void 0 : _a3.baselineOffsetEm) === void 0 ? void 0 : `${-label.style.baselineOffsetEm}em`)} data-v-52fe89f0>${ssrInterpolate(line)}</tspan>`);
            });
            _push(`<!--]--></text></g>`);
          });
          _push(`<!--]--><g class="chart-legend-group"${ssrRenderAttr("clip-path", ((__ = chartVisuals.value.get(element.id)) == null ? void 0 : __.legendClip) ? `url(#${(_aa = (_$ = chartVisuals.value.get(element.id)) == null ? void 0 : _$.legendClip) == null ? void 0 : _aa.id})` : void 0)} data-v-52fe89f0><!--[-->`);
          ssrRenderList((_ba = chartVisuals.value.get(element.id)) == null ? void 0 : _ba.legend, (entry, index) => {
            _push(`<g data-v-52fe89f0><g${ssrRenderAttr("fill", entry.color)} class="chart-legend-key"${ssrRenderAttr("data-symbol", entry.symbol)} data-v-52fe89f0>`);
            if (!entry.markerChart) _push(`<rect${ssrRenderAttr("x", entry.x)}${ssrRenderAttr("y", entry.y - 12)} width="12" height="12" data-v-52fe89f0></rect>`);
            else _push(`<!---->`);
            if (entry.line) _push(`<path${ssrRenderAttr("d", `M ${entry.x} ${entry.y - 6} H ${entry.x + 12}`)} fill="none"${ssrRenderAttr("stroke", entry.color)} stroke-width="2" stroke-linecap="round" class="chart-legend-key-line" data-v-52fe89f0></path>`);
            else _push(`<!---->`);
            if (entry.symbol) {
              _push(`<!--[-->`);
              if (!chartMarkerPath({
                x: entry.x + 6,
                y: entry.y - 6,
                color: entry.color,
                radius: 4,
                symbol: entry.symbol
              })) _push(`<circle${ssrRenderAttr("cx", entry.x + 6)}${ssrRenderAttr("cy", entry.y - 6)}${ssrRenderAttr("r", entry.symbol === "dot" ? 1.8 : 4)} class="chart-legend-key-marker"${ssrRenderAttr("data-marker-shape", entry.symbol)} data-v-52fe89f0></circle>`);
              else _push(`<path${ssrRenderAttr("d", chartMarkerPath({
                x: entry.x + 6,
                y: entry.y - 6,
                color: entry.color,
                radius: 4,
                symbol: entry.symbol
              }))} class="chart-legend-key-marker"${ssrRenderAttr("data-marker-shape", entry.symbol)} data-v-52fe89f0></path>`);
              _push(`<!--]-->`);
            } else _push(`<!---->`);
            _push(`</g><text${ssrRenderAttr("x", entry.x + 18)}${ssrRenderAttr("y", entry.y)} class="chart-label chart-legend-label" data-v-52fe89f0>${ssrInterpolate(entry.text)}</text></g>`);
          });
          _push(`<!--]--></g></svg>`);
        } else if ((element.kind === "picture" || element.kind === "video") && element.imageUrl) {
          _push(`<div class="image-viewport" aria-hidden="true" data-v-52fe89f0><!--[-->`);
          ssrRenderList(imageToneEnabled(element) ? [imageTone(element)] : [], (tone) => {
            _push(`<svg class="image-tone-defs" aria-hidden="true" width="0" height="0" data-v-52fe89f0><defs data-v-52fe89f0><filter${ssrRenderAttr("id", imageToneFilterId(element))} class="image-tone-filter" color-interpolation-filters="sRGB" data-v-52fe89f0><feComponentTransfer data-v-52fe89f0><feFuncR class="image-tone-gain" type="linear"${ssrRenderAttr("slope", tone.gain)}${ssrRenderAttr("intercept", (1 - tone.gain) / 2)} data-v-52fe89f0></feFuncR><feFuncG class="image-tone-gain" type="linear"${ssrRenderAttr("slope", tone.gain)}${ssrRenderAttr("intercept", (1 - tone.gain) / 2)} data-v-52fe89f0></feFuncG><feFuncB class="image-tone-gain" type="linear"${ssrRenderAttr("slope", tone.gain)}${ssrRenderAttr("intercept", (1 - tone.gain) / 2)} data-v-52fe89f0></feFuncB></feComponentTransfer><feComponentTransfer data-v-52fe89f0><feFuncR class="image-tone-blacklevel" type="linear" slope="1"${ssrRenderAttr("intercept", tone.blacklevel)} data-v-52fe89f0></feFuncR><feFuncG class="image-tone-blacklevel" type="linear" slope="1"${ssrRenderAttr("intercept", tone.blacklevel)} data-v-52fe89f0></feFuncG><feFuncB class="image-tone-blacklevel" type="linear" slope="1"${ssrRenderAttr("intercept", tone.blacklevel)} data-v-52fe89f0></feFuncB></feComponentTransfer><feComponentTransfer data-v-52fe89f0><feFuncR class="image-tone-gamma" type="gamma" amplitude="1"${ssrRenderAttr("exponent", 1 / tone.gamma)} offset="0" data-v-52fe89f0></feFuncR><feFuncG class="image-tone-gamma" type="gamma" amplitude="1"${ssrRenderAttr("exponent", 1 / tone.gamma)} offset="0" data-v-52fe89f0></feFuncG><feFuncB class="image-tone-gamma" type="gamma" amplitude="1"${ssrRenderAttr("exponent", 1 / tone.gamma)} offset="0" data-v-52fe89f0></feFuncB></feComponentTransfer></filter></defs></svg>`);
          });
          _push(`<!--]--><img class="slide-image"${ssrRenderAttr("src", element.imageUrl)}${ssrRenderAttr("alt", element.name)} style="${ssrRenderStyle(imageStyle(element))}" data-v-52fe89f0>`);
          if (pictureInnerShadowStyle(element)) _push(`<div class="picture-inner-shadow" style="${ssrRenderStyle(pictureInnerShadowStyle(element))}" aria-hidden="true" data-v-52fe89f0></div>`);
          else _push(`<!---->`);
          _push(`</div>`);
        } else if (element.kind === "table" && element.table) {
          _push(`<table class="slide-table" data-v-52fe89f0><colgroup data-v-52fe89f0><!--[-->`);
          ssrRenderList(element.table.columns, (column, columnIndex) => {
            _push(`<col style="${ssrRenderStyle({ width: `${column / Math.max(element.table.columns.reduce((sum, width) => sum + width, 0), 1) * 100}%` })}" data-v-52fe89f0>`);
          });
          _push(`<!--]--></colgroup><tbody data-v-52fe89f0><!--[-->`);
          ssrRenderList(element.table.rows, (row, rowIndex) => {
            _push(`<tr style="${ssrRenderStyle({ height: `${row.height / Math.max(element.table.rows.reduce((sum, item) => sum + item.height, 0), 1) * 100}%` })}" data-v-52fe89f0><!--[-->`);
            ssrRenderList(row.cells, (cell, cellIndex) => {
              _push(`<!--[-->`);
              if (!cell.hidden) {
                _push(`<td${ssrRenderAttr("colspan", cell.colSpan)}${ssrRenderAttr("rowspan", cell.rowSpan)}${ssrRenderAttr("data-autofit", cell.textAutoFitDynamic ? "normal" : void 0)}${ssrRenderAttr("data-row-index", rowIndex)}${ssrRenderAttr("data-cell-index", cellIndex)} style="${ssrRenderStyle(tableCellStyle(cell))}" data-v-52fe89f0><div class="table-cell-text" style="${ssrRenderStyle(tableCellTextStyle(cell, element.table.rows.slice(rowIndex, rowIndex + cell.rowSpan).reduce((height, item) => height + item.height, 0)))}" data-v-52fe89f0><!--[-->`);
                ssrRenderList(unref(paragraphEntries)(cell.paragraphs), (entry, paragraphIndex) => {
                  var _a2, _b2;
                  _push(`<div class="text-paragraph" style="${ssrRenderStyle(paragraphStyle(entry.paragraph, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), cell.textLineSpacingReduction, cell.textColumnsRightToLeft && ((_a2 = cell.textColumnCount) != null ? _a2 : 1) > 1 ? "ltr" : void 0))}" data-v-52fe89f0>`);
                  if (((_b2 = entry.paragraph.bullet) == null ? void 0 : _b2.kind) === "image" && entry.paragraph.bullet.url) _push(`<img class="paragraph-picture-bullet"${ssrRenderAttr("src", entry.paragraph.bullet.url)} style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" alt="" aria-hidden="true" data-v-52fe89f0>`);
                  else if (entry.marker) _push(`<span class="paragraph-marker" style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" aria-hidden="true" data-v-52fe89f0>${ssrInterpolate(entry.marker)}</span>`);
                  else _push(`<!---->`);
                  _push(`<!--[-->`);
                  ssrRenderList(entry.paragraph.runs, (run, runIndex) => {
                    _push(`<!--[--><!--[-->`);
                    ssrRenderList(runSegments(run, element, paragraphIndex, runIndex), (segment, segmentIndex) => {
                      var _a3;
                      ssrRenderVNode(_push, createVNode(resolveDynamicComponent(canActivateHyperlink(run.hyperlink) ? "a" : "span"), {
                        key: segmentIndex,
                        href: canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : void 0,
                        title: (_a3 = run.hyperlink) == null ? void 0 : _a3.tooltip,
                        "data-pptx-hyperlink": canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : void 0,
                        class: ["text-run", {
                          "text-run-hyperlink": canActivateHyperlink(run.hyperlink),
                          "text-run-break": segment.text === "\n",
                          "text-run-pattern-fill": !segment.separateDecorations && runHasTextPattern(element, run, paragraphIndex, segment.start, segment.end)
                        }],
                        "data-char-start": segment.start,
                        "data-char-end": segment.end,
                        style: [
                          runStyle(element, run, segment.script, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), paragraphIndex, segment.start, segment.end, segment.separateDecorations),
                          textColorAnimationStyle(element, paragraphIndex, segment.start, segment.end),
                          textRangeOpacityAnimationStyle(element, paragraphIndex, segment.start, segment.end),
                          characterAnimationStyle(element, paragraphIndex, segment.start, segment.end),
                          textRunAnimationStyle(element, paragraphIndex, segment.start, segment.end, run, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), segment.script, segment.separateDecorations),
                          textTransformAnimationStyle(element, paragraphIndex, segment.start, segment.end)
                        ]
                      }, {
                        default: withCtx((_, _push2, _parent2, _scopeId) => {
                          if (_push2) {
                            if (segment.separateDecorations) _push2(`<span class="${ssrRenderClass([{ "text-run-pattern-fill": runHasTextPattern(element, run, paragraphIndex, segment.start, segment.end) }, "text-run-underline"])}" style="${ssrRenderStyle(textUnderlineStyle(run))}" data-v-52fe89f0${_scopeId}>${ssrInterpolate(segment.text)}</span>`);
                            else _push2(`<!--[-->${ssrInterpolate(segment.text)}<!--]-->`);
                          } else return [segment.separateDecorations ? (openBlock(), createBlock("span", {
                            key: 0,
                            class: ["text-run-underline", { "text-run-pattern-fill": runHasTextPattern(element, run, paragraphIndex, segment.start, segment.end) }],
                            style: textUnderlineStyle(run)
                          }, toDisplayString(segment.text), 7)) : (openBlock(), createBlock(Fragment, { key: 1 }, [createTextVNode(toDisplayString(segment.text), 1)], 64))];
                        }),
                        _: 2
                      }), _parent);
                    });
                    _push(`<!--]--><!--]-->`);
                  });
                  _push(`<!--]--></div>`);
                });
                _push(`<!--]--></div></td>`);
              } else _push(`<!---->`);
              _push(`<!--]-->`);
            });
            _push(`<!--]--></tr>`);
          });
          _push(`<!--]--></tbody></table>`);
        } else if (element.kind === "placeholder") _push(`<span data-v-52fe89f0>${ssrInterpolate(element.mediaType ? "Media unavailable" : "Image unavailable")}</span>`);
        else _push(`<!---->`);
        if (element.paragraphs.length) {
          _push(`<div class="text-frame"${ssrRenderAttr("data-element-id", element.id)}${ssrRenderAttr("data-autofit", element.textAutoFit === "shape" ? "shape" : element.textAutoFitDynamic ? "normal" : void 0)}${ssrRenderAttr("data-text-overflow", supportsTextVerticalEllipsis(element) ? "ellipsis" : void 0)} style="${ssrRenderStyle(textFrameStyle(element))}" data-v-52fe89f0><!--[-->`);
          ssrRenderList(unref(paragraphEntries)(element.paragraphs), (entry, index) => {
            var _a2;
            _push(`<div class="text-paragraph" style="${ssrRenderStyle([paragraphStyle(entry.paragraph, effectiveTextFontScale(element), element.textLineSpacingReduction, usesRightToLeftColumns(element) ? "ltr" : void 0), paragraphAnimationStyle(element, index)])}" data-v-52fe89f0>`);
            if (((_a2 = entry.paragraph.bullet) == null ? void 0 : _a2.kind) === "image" && entry.paragraph.bullet.url) _push(`<img class="paragraph-picture-bullet"${ssrRenderAttr("src", entry.paragraph.bullet.url)} style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" alt="" aria-hidden="true" data-v-52fe89f0>`);
            else if (entry.marker) _push(`<span class="paragraph-marker" style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" aria-hidden="true" data-v-52fe89f0>${ssrInterpolate(entry.marker)}</span>`);
            else _push(`<!---->`);
            _push(`<!--[-->`);
            ssrRenderList(entry.paragraph.runs, (run, runIndex) => {
              _push(`<!--[--><!--[-->`);
              ssrRenderList(runSegments(run, element, index, runIndex), (segment, segmentIndex) => {
                var _a3;
                ssrRenderVNode(_push, createVNode(resolveDynamicComponent(canActivateHyperlink(run.hyperlink) ? "a" : "span"), {
                  key: segmentIndex,
                  href: canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : void 0,
                  title: (_a3 = run.hyperlink) == null ? void 0 : _a3.tooltip,
                  "data-pptx-hyperlink": canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : void 0,
                  class: ["text-run", {
                    "text-run-hyperlink": canActivateHyperlink(run.hyperlink),
                    "text-run-break": segment.text === "\n",
                    "text-run-pattern-fill": !segment.separateDecorations && runHasTextPattern(element, run, index, segment.start, segment.end)
                  }],
                  "data-char-start": segment.start,
                  "data-char-end": segment.end,
                  style: [
                    runStyle(element, run, segment.script, effectiveTextFontScale(element), index, segment.start, segment.end, segment.separateDecorations),
                    textColorAnimationStyle(element, index, segment.start, segment.end),
                    textRangeOpacityAnimationStyle(element, index, segment.start, segment.end),
                    characterAnimationStyle(element, index, segment.start, segment.end),
                    textRunAnimationStyle(element, index, segment.start, segment.end, run, effectiveTextFontScale(element), segment.script, segment.separateDecorations),
                    textTransformAnimationStyle(element, index, segment.start, segment.end)
                  ]
                }, {
                  default: withCtx((_, _push2, _parent2, _scopeId) => {
                    if (_push2) {
                      if (segment.separateDecorations) _push2(`<span class="${ssrRenderClass([{ "text-run-pattern-fill": runHasTextPattern(element, run, index, segment.start, segment.end) }, "text-run-underline"])}" style="${ssrRenderStyle(textUnderlineStyle(run))}" data-v-52fe89f0${_scopeId}>${ssrInterpolate(segment.text)}</span>`);
                      else _push2(`<!--[-->${ssrInterpolate(segment.text)}<!--]-->`);
                    } else return [segment.separateDecorations ? (openBlock(), createBlock("span", {
                      key: 0,
                      class: ["text-run-underline", { "text-run-pattern-fill": runHasTextPattern(element, run, index, segment.start, segment.end) }],
                      style: textUnderlineStyle(run)
                    }, toDisplayString(segment.text), 7)) : (openBlock(), createBlock(Fragment, { key: 1 }, [createTextVNode(toDisplayString(segment.text), 1)], 64))];
                  }),
                  _: 2
                }), _parent);
              });
              _push(`<!--]--><!--]-->`);
            });
            _push(`<!--]--></div>`);
          });
          _push(`<!--]--></div>`);
        } else _push(`<!---->`);
        if (element.kind === "line") {
          _push(`<svg class="line-render" style="${ssrRenderStyle(lineRenderStyle(element))}" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true" data-v-52fe89f0><defs data-v-52fe89f0>`);
          if (animationLineEnd(element, "head").type !== "none") _push(`<marker${ssrRenderAttr("id", lineEndId(element, "head"))}${ssrRenderAttr("markerWidth", lineEndMarkerSize(element, "head").width)}${ssrRenderAttr("markerHeight", lineEndMarkerSize(element, "head").height)} viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto-start-reverse" data-v-52fe89f0><path${ssrRenderAttr("d", lineEndPath(animationLineEnd(element, "head").type))}${ssrRenderAttr("fill", lineEndUsesStroke(animationLineEnd(element, "head").type) ? "none" : "currentColor")}${ssrRenderAttr("stroke", lineEndUsesStroke(animationLineEnd(element, "head").type) ? "currentColor" : "none")} data-v-52fe89f0></path></marker>`);
          else _push(`<!---->`);
          if (animationLineEnd(element, "tail").type !== "none") _push(`<marker${ssrRenderAttr("id", lineEndId(element, "tail"))}${ssrRenderAttr("markerWidth", lineEndMarkerSize(element, "tail").width)}${ssrRenderAttr("markerHeight", lineEndMarkerSize(element, "tail").height)} viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto" data-v-52fe89f0><path${ssrRenderAttr("d", lineEndPath(animationLineEnd(element, "tail").type))}${ssrRenderAttr("fill", lineEndUsesStroke(animationLineEnd(element, "tail").type) ? "none" : "currentColor")}${ssrRenderAttr("stroke", lineEndUsesStroke(animationLineEnd(element, "tail").type) ? "currentColor" : "none")} data-v-52fe89f0></path></marker>`);
          else _push(`<!---->`);
          _push(`</defs>`);
          if (element.linePath) _push(`<path${ssrRenderAttr("d", element.linePath)} fill="none"${ssrRenderAttr("stroke", lineStrokeColor(element))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}"${ssrRenderAttr("marker-start", lineEndUrl(element, "head"))}${ssrRenderAttr("marker-end", lineEndUrl(element, "tail"))} vector-effect="non-scaling-stroke" data-v-52fe89f0></path>`);
          else _push(`<line x1="0"${ssrRenderAttr("y1", element.geometry === "lineInv" ? 1e3 : 0)}${ssrRenderAttr("x2", element.width > 0 ? 1e3 : 0)}${ssrRenderAttr("y2", element.height > 0 ? element.geometry === "lineInv" ? 0 : 1e3 : 0)}${ssrRenderAttr("stroke", lineStrokeColor(element))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}"${ssrRenderAttr("marker-start", lineEndUrl(element, "head"))}${ssrRenderAttr("marker-end", lineEndUrl(element, "tail"))} vector-effect="non-scaling-stroke" data-v-52fe89f0></line>`);
          _push(`</svg>`);
        } else _push(`<!---->`);
        _push(`</div>`);
      });
      _push(`<!--]-->`);
      if (backgroundTileWarning.value) _push(`<span class="media-warning" role="status" data-v-52fe89f0>${ssrInterpolate(backgroundTileWarning.value)}</span>`);
      else _push(`<!---->`);
      if (textOverflowWarning.value && (!__props.thumbnail || __props.playbackPreview)) _push(`<span class="media-warning" role="status" data-v-52fe89f0>${ssrInterpolate(textOverflowWarning.value)}</span>`);
      else _push(`<!---->`);
      _push(`</div>`);
    };
  }
});
var _sfc_setup = SlideStage_vue_vue_type_script_setup_true_lang_default.setup;
SlideStage_vue_vue_type_script_setup_true_lang_default.setup = (props, ctx) => {
  const ssrContext = useSSRContext();
  (ssrContext.modules || (ssrContext.modules = /* @__PURE__ */ new Set())).add("components/SlideStage.vue");
  return _sfc_setup ? _sfc_setup(props, ctx) : void 0;
};
var SlideStage_default = /* @__PURE__ */ Object.assign(_plugin_vue_export_helper_default(SlideStage_vue_vue_type_script_setup_true_lang_default, [["__scopeId", "data-v-52fe89f0"]]), { __name: "SlideStage" });
function honeycombClipPath(width, height, collapsed) {
  const side = Math.sqrt(width * height / 340);
  const halfWidth = Math.sqrt(3) * side / 2;
  const rowStep = side * 1.5;
  const columnStep = halfWidth * 2;
  const vertices = [
    [0, -side],
    [halfWidth, -side / 2],
    [halfWidth, side / 2],
    [0, side],
    [-halfWidth, side / 2],
    [-halfWidth, -side / 2]
  ];
  const paths = [];
  for (let row = -Math.ceil(height / rowStep) - 2; row <= Math.ceil(height / rowStep) + 2; row++) for (let column = -Math.ceil(width / columnStep) - 2; column <= Math.ceil(width / columnStep) + 2; column++) {
    const x = width / 2 + columnStep * (column + row / 2);
    const y = height / 2 + rowStep * row;
    if (x + halfWidth < 0 || x - halfWidth > width || y + side < 0 || y - side > height) continue;
    paths.push(`M${x.toFixed(2)} ${y.toFixed(2)} ${vertices.map(([dx, dy]) => `L${(x + (collapsed ? 0 : dx)).toFixed(2)} ${(y + (collapsed ? 0 : dy)).toFixed(2)}`).join(" ")} Z`);
  }
  return `path("${paths.join(" ")}")`;
}
function morphTextSegments(text, mode) {
  if (typeof Intl.Segmenter !== "function") {
    if (mode === "byChar") {
      let offset = 0;
      return Array.from(text, (segment) => {
        const start = offset;
        offset += segment.length;
        return {
          text: segment,
          key: segment,
          start,
          end: offset
        };
      });
    }
    return [...text.matchAll(/\S+\s*|\s+/gu)].map((match) => ({
      text: match[0],
      key: match[0].trimEnd() || match[0],
      start: match.index,
      end: match.index + match[0].length
    }));
  }
  const segments = [...new Intl.Segmenter(void 0, { granularity: mode === "byChar" ? "grapheme" : "word" }).segment(text)];
  if (mode === "byChar") return segments.map((segment) => ({
    text: segment.segment,
    key: segment.segment,
    start: segment.index,
    end: segment.index + segment.segment.length
  }));
  const words = segments.filter((segment) => segment.isWordLike);
  if (!words.length) return text ? [{
    text,
    key: text,
    start: 0,
    end: text.length
  }] : [];
  return words.map((word, index) => {
    const start = index === 0 ? 0 : word.index;
    const end = index + 1 < words.length ? words[index + 1].index : text.length;
    return {
      text: text.slice(start, end),
      key: word.segment,
      start,
      end
    };
  });
}
function morphTextTokens(frame, mode) {
  const tokens = [];
  const frameRect = frame.getBoundingClientRect();
  for (const paragraph of frame.querySelectorAll(".text-paragraph")) {
    const nodes = [];
    let text = "";
    for (const run of paragraph.querySelectorAll(".text-run")) {
      const walker = (void 0).createTreeWalker(run, NodeFilter.SHOW_TEXT);
      let node = walker.nextNode();
      while (node) {
        const start = text.length;
        text += node.data;
        nodes.push({
          node,
          start,
          end: text.length
        });
        node = walker.nextNode();
      }
    }
    for (const segment of morphTextSegments(text, mode)) {
      const start = nodes.find((item) => segment.start >= item.start && segment.start < item.end);
      const end = [...nodes].reverse().find((item) => segment.end > item.start && segment.end <= item.end);
      if (!start || !end) continue;
      const range = (void 0).createRange();
      range.setStart(start.node, segment.start - start.start);
      range.setEnd(end.node, segment.end - end.start);
      const rect = range.getBoundingClientRect();
      if (!rect.width || !rect.height) continue;
      const computed2 = getComputedStyle(start.node.parentElement || frame);
      tokens.push({
        text: segment.text,
        key: segment.key,
        left: rect.left - frameRect.left,
        top: rect.top - frameRect.top,
        width: rect.width,
        height: rect.height,
        content: range.cloneContents(),
        style: {
          color: computed2.color,
          fontFamily: computed2.fontFamily,
          fontSize: computed2.fontSize,
          fontStyle: computed2.fontStyle,
          fontWeight: computed2.fontWeight,
          fontVariant: computed2.fontVariant,
          fontStretch: computed2.fontStretch,
          lineHeight: computed2.lineHeight,
          letterSpacing: computed2.letterSpacing,
          wordSpacing: computed2.wordSpacing,
          textDecoration: computed2.textDecoration,
          textShadow: computed2.textShadow,
          textTransform: computed2.textTransform,
          direction: computed2.direction,
          writingMode: computed2.writingMode,
          verticalAlign: computed2.verticalAlign
        }
      });
    }
  }
  return tokens;
}
function morphTextNode(layer, token, left, top) {
  const node = (void 0).createElement("span");
  node.className = "morph-text-token";
  node.dataset.morphToken = token.text;
  node.dataset.morphMatch = token.key;
  node.setAttribute("aria-hidden", "true");
  Object.assign(node.style, {
    position: "absolute",
    display: "block",
    left: `${left}px`,
    top: `${top}px`,
    width: `${Math.max(token.width, 0.01)}px`,
    height: `${Math.max(token.height, 0.01)}px`,
    margin: "0",
    padding: "0",
    border: "0",
    overflow: "visible",
    whiteSpace: "pre",
    pointerEvents: "none",
    transformOrigin: "top left",
    ...token.style
  });
  node.append(token.content);
  for (const run of node.querySelectorAll(".text-run")) run.style.setProperty("white-space", "pre", "important");
  layer.append(node);
  return node;
}
function useSlideTransition({ slide, shell, incomingHost, underlay, flash, playback }) {
  let morphCleanup;
  watch(slide, (nextSlide, previousSlide) => {
    var _a, _b;
    if ((nextSlide == null ? void 0 : nextSlide.id) === (previousSlide == null ? void 0 : previousSlide.id)) return;
    morphCleanup == null ? void 0 : morphCleanup();
    morphCleanup = void 0;
    (_a = underlay.value) == null ? void 0 : _a.replaceChildren();
    if (!nextSlide) return;
    const current = incomingHost.value;
    if (!current) return;
    const snapshot = current.cloneNode(true);
    const style = getComputedStyle(current);
    snapshot.setAttribute("aria-hidden", "true");
    snapshot.inert = true;
    snapshot.style.pointerEvents = "none";
    snapshot.style.transform = style.transform;
    snapshot.style.opacity = style.opacity;
    snapshot.style.clipPath = style.clipPath;
    (_b = underlay.value) == null ? void 0 : _b.append(snapshot);
  }, { flush: "sync" });
  watch(slide, async (activeSlide, previousSlide) => {
    var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y, _z, _A, _B, _C, _D;
    if (!activeSlide || activeSlide.id === (previousSlide == null ? void 0 : previousSlide.id)) return;
    const currentPlayback = ((_a = playback.value) == null ? void 0 : _a.slideId) === activeSlide.id ? playback.value : null;
    const transition = currentPlayback == null ? void 0 : currentPlayback.transition;
    await nextTick();
    if (((_b = slide.value) == null ? void 0 : _b.id) !== activeSlide.id) return;
    const animationDelay = currentPlayback ? Math.min(0, currentPlayback.startedAtEpoch - (performance.timeOrigin + performance.now())) : 0;
    (_c = shell.value) == null ? void 0 : _c.classList.remove("is-fading-through-black", "is-cut-through-black", "is-prism-transition", "is-flip-transition", "is-switch-transition", "is-ferris-transition", "is-gallery-transition", "is-conveyor-transition", "is-morph-transition", "is-preset-transition");
    (_d = underlay.value) == null ? void 0 : _d.classList.remove("is-shredding-out", "is-flythrough", "is-warping-out", "is-vortexing");
    const incoming = incomingHost.value;
    incoming == null ? void 0 : incoming.getAnimations().forEach((animation2) => animation2.cancel());
    (_e = flash.value) == null ? void 0 : _e.getAnimations().forEach((animation2) => animation2.cancel());
    if (!transition || (void 0).matchMedia("(prefers-reduced-motion: reduce)").matches || !incoming || !((_f = underlay.value) == null ? void 0 : _f.firstElementChild)) {
      (_g = underlay.value) == null ? void 0 : _g.replaceChildren();
      return;
    }
    if (transition.effect === "morph") {
      const outgoing2 = (_h = underlay.value) == null ? void 0 : _h.firstElementChild;
      const stage = shell.value;
      if (!outgoing2 || !stage) {
        (_i = underlay.value) == null ? void 0 : _i.replaceChildren();
        return;
      }
      const timing2 = {
        duration: transition.durationMs,
        easing: "ease-in-out",
        fill: "both",
        delay: animationDelay
      };
      const animations = [];
      const textMode = transition.morphOption === "byWord" || transition.morphOption === "byChar" ? transition.morphOption : void 0;
      stage.classList.add("is-morph-transition");
      const outgoingFrame = outgoing2.querySelector(".slide-frame");
      const incomingFrame = incoming.querySelector(".slide-frame");
      const outgoingElements = [...(outgoingFrame == null ? void 0 : outgoingFrame.querySelectorAll(".slide-element")) || []];
      const incomingElements = [...(incomingFrame == null ? void 0 : incomingFrame.querySelectorAll(".slide-element")) || []];
      const uniqueByName = (elements) => {
        const grouped = /* @__PURE__ */ new Map();
        for (const element of elements) {
          const name = element.dataset.elementName;
          const kind = element.dataset.elementKind;
          if (!name || !kind) continue;
          const key = `${kind}:${name}`;
          grouped.set(key, [...grouped.get(key) || [], element]);
        }
        return new Map([...grouped].filter(([, matches]) => matches.length === 1).map(([key, matches]) => [key, matches[0]]));
      };
      const oldByName = uniqueByName(outgoingElements);
      const newByName = uniqueByName(incomingElements);
      const restoreTextStyles = [];
      if (textMode && incomingFrame) {
        const transformed = (element) => {
          const value = getComputedStyle(element).transform;
          if (value === "none") return false;
          try {
            const matrix = new DOMMatrixReadOnly(value);
            return Math.abs(matrix.a - 1) > 1e-3 || Math.abs(matrix.b) > 1e-3 || Math.abs(matrix.c) > 1e-3 || Math.abs(matrix.d - 1) > 1e-3;
          } catch {
            return true;
          }
        };
        for (const [key, element] of newByName) {
          const previous = oldByName.get(key);
          if (!previous) continue;
          const sourceText = previous.querySelector(".text-frame");
          const targetText = element.querySelector(".text-frame");
          if (!sourceText || !targetText) continue;
          if (element.dataset.elementKind !== "shape") {
            console.warn(`Morph ${textMode} text matching is not supported on ${element.dataset.elementKind} elements; object motion is used.`);
            continue;
          }
          if (transformed(previous) || transformed(element)) {
            console.warn(`Morph ${textMode} text matching uses object motion for transformed shape "${element.dataset.elementName}".`);
            continue;
          }
          const sourceTokens = morphTextTokens(sourceText, textMode);
          const targetTokens = morphTextTokens(targetText, textMode);
          if (!sourceTokens.length && !targetTokens.length) continue;
          const position = targetText.style.getPropertyValue("position");
          const positionPriority = targetText.style.getPropertyPriority("position");
          targetText.style.setProperty("position", "relative");
          restoreTextStyles.push(() => position ? targetText.style.setProperty("position", position, positionPriority) : targetText.style.removeProperty("position"));
          const layer = (void 0).createElement("span");
          layer.className = "morph-text-layer";
          layer.setAttribute("aria-hidden", "true");
          Object.assign(layer.style, {
            position: "absolute",
            display: "block",
            inset: "0",
            overflow: "hidden",
            pointerEvents: "none",
            zIndex: "2"
          });
          targetText.append(layer);
          const available = /* @__PURE__ */ new Map();
          for (const token of targetTokens) {
            const bucket = available.get(token.key) || {
              tokens: [],
              index: 0
            };
            bucket.tokens.push(token);
            available.set(token.key, bucket);
          }
          const matchedTargets = /* @__PURE__ */ new Set();
          for (const token of sourceTokens) {
            const matches = available.get(token.key);
            if (matches && matches.index < matches.tokens.length) {
              const target = matches.tokens[matches.index++];
              matchedTargets.add(target);
              const node = morphTextNode(layer, target, token.left, token.top);
              const scaleX = token.width / Math.max(target.width, 0.01);
              const scaleY = token.height / Math.max(target.height, 0.01);
              animations.push(node.animate([{
                left: `${token.left}px`,
                top: `${token.top}px`,
                transform: `scale(${scaleX}, ${scaleY})`
              }, {
                left: `${target.left}px`,
                top: `${target.top}px`,
                transform: "scale(1, 1)"
              }], timing2));
            } else {
              const node = morphTextNode(layer, token, token.left, token.top);
              animations.push(node.animate([{ opacity: 1 }, { opacity: 0 }], timing2));
            }
          }
          for (const token of targetTokens) {
            if (matchedTargets.has(token)) continue;
            const node = morphTextNode(layer, token, token.left, token.top);
            animations.push(node.animate([{ opacity: 0 }, { opacity: 1 }], timing2));
          }
          for (const run of sourceText.querySelectorAll(".text-run")) run.style.setProperty("visibility", "hidden", "important");
          for (const run of targetText.querySelectorAll(".text-run")) {
            const visibility = run.style.getPropertyValue("visibility");
            const priority = run.style.getPropertyPriority("visibility");
            run.style.setProperty("visibility", "hidden", "important");
            restoreTextStyles.push(() => visibility ? run.style.setProperty("visibility", visibility, priority) : run.style.removeProperty("visibility"));
          }
        }
      }
      for (const element of incomingElements) {
        const key = `${element.dataset.elementKind}:${element.dataset.elementName}`;
        const previous = oldByName.get(key);
        const finalOpacity = getComputedStyle(element).opacity;
        if (!previous || !newByName.has(key)) {
          animations.push(element.animate([{ opacity: 0 }, { opacity: finalOpacity }], timing2));
          continue;
        }
        const color = (node, property) => getComputedStyle(node)[property];
        animations.push(element.animate([{
          left: previous.style.left,
          top: previous.style.top,
          width: previous.style.width,
          height: previous.style.height,
          transform: previous.style.transform || "none",
          transformOrigin: previous.style.transformOrigin || "50% 50%",
          opacity: getComputedStyle(previous).opacity,
          backgroundColor: color(previous, "backgroundColor"),
          borderColor: color(previous, "borderColor"),
          borderWidth: color(previous, "borderWidth")
        }, {
          left: element.style.left,
          top: element.style.top,
          width: element.style.width,
          height: element.style.height,
          transform: element.style.transform || "none",
          transformOrigin: element.style.transformOrigin || "50% 50%",
          opacity: finalOpacity,
          backgroundColor: color(element, "backgroundColor"),
          borderColor: color(element, "borderColor"),
          borderWidth: color(element, "borderWidth")
        }], timing2));
      }
      for (const element of outgoingElements) animations.push(element.animate([{ opacity: getComputedStyle(element).opacity }, { opacity: 0 }], timing2));
      if (incomingFrame) {
        const backgroundAnimation = incomingFrame.animate([{ backgroundColor: "rgba(0, 0, 0, 0)" }, { backgroundColor: getComputedStyle(incomingFrame).backgroundColor }], timing2);
        animations.push(backgroundAnimation);
        let cleaned = false;
        const cleanup = () => {
          if (cleaned) return;
          cleaned = true;
          animations.forEach((animation2) => animation2.cancel());
          restoreTextStyles.forEach((restore) => restore());
          incomingFrame.querySelectorAll(".morph-text-layer").forEach((layer) => layer.remove());
          if (morphCleanup === cleanup) morphCleanup = void 0;
        };
        morphCleanup = cleanup;
        backgroundAnimation.onfinish = () => {
          var _a2, _b2;
          if (((_a2 = slide.value) == null ? void 0 : _a2.id) === activeSlide.id) {
            cleanup();
            (_b2 = underlay.value) == null ? void 0 : _b2.replaceChildren();
            stage.classList.remove("is-morph-transition");
          }
        };
      }
      return;
    }
    if (transition.effect === "preset") {
      const outgoing2 = (_j = underlay.value) == null ? void 0 : _j.firstElementChild;
      const stage = shell.value;
      if (!outgoing2 || !stage || !transition.presetName) {
        (_k = underlay.value) == null ? void 0 : _k.replaceChildren();
        return;
      }
      const x = transition.invertX ? -1 : 1;
      const y = transition.invertY ? -1 : 1;
      const origin = x > 0 ? "left center" : "right center";
      const frames2 = (() => {
        switch (transition.presetName) {
          case "fallOver":
            return [{
              transform: "perspective(1200px) rotateX(0) rotateZ(0)",
              opacity: 1
            }, {
              transform: `perspective(1200px) rotateX(${72 * y}deg) rotateZ(${5 * x}deg) translateY(${10 * y}%)`,
              opacity: 0
            }];
          case "drape":
            return [
              {
                transform: "perspective(1200px) rotateX(0) scaleY(1)",
                opacity: 1
              },
              {
                transform: `perspective(1200px) rotateX(${48 * y}deg) scaleY(.9)`,
                opacity: 0.7,
                offset: 0.55
              },
              {
                transform: `perspective(1200px) rotateX(${78 * y}deg) scaleY(.7)`,
                opacity: 0
              }
            ];
          case "curtains":
            return [{
              transform: "scaleX(1) rotateY(0)",
              transformOrigin: origin,
              opacity: 1
            }, {
              transform: `scaleX(.02) rotateY(${8 * y}deg)`,
              transformOrigin: origin,
              opacity: 0
            }];
          case "wind":
            return [{
              transform: "translate(0, 0) skewX(0)",
              opacity: 1
            }, {
              transform: `translate(${125 * x}%, ${ -110 * y}%) skewX(${14 * y}deg)`,
              opacity: 0
            }];
          case "prestige":
            return [{
              transform: "scale(1) rotate(0)",
              opacity: 1
            }, {
              transform: `scale(.62) rotate(${5 * x}deg) translate(${5 * x}%, ${12 * y}%)`,
              opacity: 0
            }];
          case "fracture":
            return [{
              clipPath: "inset(0)",
              transform: "scale(1) rotate(0)",
              opacity: 1
            }, {
              clipPath: `polygon(50% ${50 - 4 * y}%, ${50 + 4 * x}% 50%, 50% ${50 + 4 * y}%, ${50 - 4 * x}% 50%)`,
              transform: `scale(.82) rotate(${8 * x}deg)`,
              opacity: 0
            }];
          case "crush":
            return [{
              transform: "scale(1)",
              opacity: 1
            }, {
              transform: "scaleX(.04)",
              opacity: 0
            }];
          case "peelOff":
            return [{
              transform: "perspective(1200px) rotateY(0)",
              transformOrigin: origin,
              opacity: 1
            }, {
              transform: `perspective(1200px) rotateY(${105 * x}deg)`,
              transformOrigin: origin,
              opacity: 0
            }];
          case "pageCurlDouble":
            return [
              {
                transform: "perspective(1200px) rotateY(0) scaleX(1)",
                transformOrigin: "center center",
                opacity: 1
              },
              {
                transform: `perspective(1200px) rotateY(${82 * x}deg) scaleX(.28)`,
                transformOrigin: "center center",
                opacity: 0.8,
                offset: 0.55
              },
              {
                transform: `perspective(1200px) rotateY(${165 * x}deg) scaleX(.04)`,
                transformOrigin: "center center",
                opacity: 0
              }
            ];
          case "pageCurlSingle":
            return [{
              transform: "perspective(1200px) rotateY(0)",
              transformOrigin: origin,
              opacity: 1
            }, {
              transform: `perspective(1200px) rotateY(${112 * x}deg)`,
              transformOrigin: origin,
              opacity: 0
            }];
          case "airplane":
            return [{
              transform: "perspective(1200px) translate(0, 0) rotate(0) rotateY(0) scale(1)",
              opacity: 1
            }, {
              transform: `perspective(1200px) translate(${130 * x}%, ${95 * y}%) rotate(${18 * x}deg) rotateY(${42 * y}deg) scale(.72)`,
              opacity: 0
            }];
          case "origami":
            return [
              {
                transform: "perspective(1200px) rotateX(0) rotateY(0) scale(1)",
                opacity: 1
              },
              {
                transform: `perspective(1200px) rotateX(${24 * y}deg) rotateY(${34 * x}deg) scale(.88)`,
                opacity: 0.9,
                offset: 0.32
              },
              {
                transform: `perspective(1200px) rotateX(${ -32 * y}deg) rotateY(${ -42 * x}deg) scale(.68)`,
                opacity: 0.65,
                offset: 0.68
              },
              {
                transform: `perspective(1200px) rotateX(${90 * y}deg) rotateY(${100 * x}deg) translate(${30 * x}%, ${30 * y}%) scale(.48)`,
                opacity: 0
              }
            ];
        }
      })();
      stage.classList.add("is-preset-transition");
      const animation2 = outgoing2.animate(frames2, {
        duration: transition.durationMs,
        easing: "ease-in-out",
        fill: "both",
        delay: animationDelay
      });
      animation2.onfinish = () => {
        var _a2, _b2;
        animation2.cancel();
        if (((_a2 = slide.value) == null ? void 0 : _a2.id) === activeSlide.id) {
          (_b2 = underlay.value) == null ? void 0 : _b2.replaceChildren();
          stage.classList.remove("is-preset-transition");
        }
      };
      return;
    }
    if (transition.effect === "cut") {
      (_l = underlay.value) == null ? void 0 : _l.replaceChildren();
      if (transition.throughBlack) {
        const stage = shell.value;
        stage == null ? void 0 : stage.classList.add("is-cut-through-black");
        requestAnimationFrame(() => requestAnimationFrame(() => {
          var _a2;
          if (((_a2 = slide.value) == null ? void 0 : _a2.id) === activeSlide.id) stage == null ? void 0 : stage.classList.remove("is-cut-through-black");
        }));
      }
      return;
    }
    const direction = transition.direction;
    const effect = transition.effect;
    if ([
      "prism",
      "flip",
      "switch",
      "ferris",
      "gallery",
      "conveyor"
    ].includes(effect)) {
      const outgoing2 = (_m = underlay.value) == null ? void 0 : _m.firstElementChild;
      const stage = shell.value;
      if (!outgoing2 || !stage) {
        (_n = underlay.value) == null ? void 0 : _n.replaceChildren();
        return;
      }
      const rotations = {
        l: {
          axis: "0, 1, 0",
          angle: 90
        },
        r: {
          axis: "0, 1, 0",
          angle: -90
        },
        u: {
          axis: "1, 0, 0",
          angle: -90
        },
        d: {
          axis: "1, 0, 0",
          angle: 90
        },
        ld: {
          axis: "1, 1, 0",
          angle: 90
        },
        lu: {
          axis: "1, -1, 0",
          angle: -90
        },
        rd: {
          axis: "1, -1, 0",
          angle: 90
        },
        ru: {
          axis: "1, 1, 0",
          angle: -90
        }
      };
      const rotation = effect === "prism" ? rotations[direction] || rotations.l : {
        axis: "0, 1, 0",
        angle: (direction === "r" ? 1 : -1) * (effect === "flip" ? 180 : 90)
      };
      const angle = rotation.angle * (effect === "prism" && transition.isInverted ? -1 : 1);
      const incomingBase = incoming.style.transform || "none";
      const outgoingBase = outgoing2.style.transform || "none";
      const base = (transform) => transform === "none" ? "" : `${transform} `;
      const side = direction === "r" ? -1 : 1;
      const startTransform = effect === "ferris" ? `${base(incomingBase)}perspective(1200px) rotateZ(${side * 180}deg) scale(0.78)` : effect === "gallery" ? `${base(incomingBase)}translateX(${side * 36}%) perspective(1200px) rotateY(${side * 38}deg) scale(0.84)` : effect === "conveyor" ? `${base(incomingBase)}translateX(${side * 100}%) perspective(1200px) rotateX(-10deg) rotateY(${side * 12}deg)` : `${base(incomingBase)}rotate3d(${rotation.axis}, ${angle}deg)`;
      const endTransform = effect === "ferris" ? `${base(outgoingBase)}perspective(1200px) rotateZ(${side * -180}deg) scale(0.78)` : effect === "gallery" ? `${base(outgoingBase)}translateX(${side * -36}%) perspective(1200px) rotateY(${side * -38}deg) scale(0.84)` : effect === "conveyor" ? `${base(outgoingBase)}translateX(${side * -100}%) perspective(1200px) rotateX(10deg) rotateY(${side * -12}deg)` : `${base(outgoingBase)}rotate3d(${rotation.axis}, ${-angle}deg)`;
      const transitionClass = effect === "prism" ? "is-prism-transition" : effect === "flip" ? "is-flip-transition" : effect === "switch" ? "is-switch-transition" : `is-${effect}-transition`;
      stage.classList.add(transitionClass);
      const timing2 = {
        duration: transition.durationMs,
        easing: "ease-in-out",
        fill: "both",
        delay: animationDelay
      };
      const incomingAnimation = incoming.animate([{ transform: startTransform }, { transform: incomingBase }], timing2);
      const outgoingAnimation2 = outgoing2.animate([{ transform: outgoingBase }, { transform: endTransform }], timing2);
      incomingAnimation.onfinish = () => {
        var _a2, _b2;
        incomingAnimation.cancel();
        outgoingAnimation2.cancel();
        if (((_a2 = slide.value) == null ? void 0 : _a2.id) === activeSlide.id) {
          (_b2 = underlay.value) == null ? void 0 : _b2.replaceChildren();
          stage.classList.remove(transitionClass);
        }
      };
      return;
    }
    if (effect === "flash") {
      const animation2 = (_o = flash.value) == null ? void 0 : _o.animate([
        { opacity: 0 },
        {
          opacity: 1,
          offset: 0.12
        },
        {
          opacity: 1,
          offset: 0.28
        },
        { opacity: 0 }
      ], {
        duration: transition.durationMs,
        easing: "linear",
        fill: "both",
        delay: animationDelay
      });
      if (animation2) animation2.onfinish = () => animation2.cancel();
      (_p = underlay.value) == null ? void 0 : _p.replaceChildren();
      return;
    }
    const revealThroughBlack = effect === "reveal" && transition.throughBlack;
    let frames;
    let outgoingTransform = "";
    let outgoingFrames;
    let shredOutgoing = false;
    if (effect === "fade") frames = transition.throughBlack ? [
      {
        opacity: 0,
        offset: 0
      },
      {
        opacity: 0,
        offset: 0.5
      },
      {
        opacity: 1,
        offset: 1
      }
    ] : [{ opacity: 0 }, { opacity: 1 }];
    else if (effect === "cover") {
      const starts = {
        l: "translateX(100%)",
        r: "translateX(-100%)",
        u: "translateY(100%)",
        d: "translateY(-100%)",
        ld: "translate(100%, -100%)",
        lu: "translate(100%, 100%)",
        rd: "translate(-100%, -100%)",
        ru: "translate(-100%, 100%)"
      };
      frames = [{ transform: starts[direction] || starts.l }, { transform: "translate(0, 0)" }];
    } else if (effect === "push" || effect === "reveal" || effect === "pan") {
      const panStarts = {
        l: "translateX(100%)",
        r: "translateX(-100%)",
        u: "translateY(100%)",
        d: "translateY(-100%)",
        ld: "translate(100%, -100%)",
        lu: "translate(100%, 100%)",
        rd: "translate(-100%, -100%)",
        ru: "translate(-100%, 100%)"
      };
      const panExits = {
        l: "translateX(-100%)",
        r: "translateX(100%)",
        u: "translateY(-100%)",
        d: "translateY(100%)",
        ld: "translate(-100%, 100%)",
        lu: "translate(-100%, -100%)",
        rd: "translate(100%, 100%)",
        ru: "translate(100%, -100%)"
      };
      const start = effect === "pan" ? panStarts[direction] || panStarts.l : direction === "u" ? "translateY(100%)" : direction === "d" ? "translateY(-100%)" : direction === "r" ? "translateX(-100%)" : "translateX(100%)";
      outgoingTransform = effect === "pan" ? panExits[direction] || panExits.l : direction === "u" ? "translateY(-100%)" : direction === "d" ? "translateY(100%)" : direction === "r" ? "translateX(100%)" : "translateX(-100%)";
      frames = revealThroughBlack ? [
        {
          transform: start,
          opacity: 0
        },
        {
          transform: start,
          opacity: 0,
          offset: 0.5
        },
        {
          transform: "translate(0, 0)",
          opacity: 1
        }
      ] : [{ transform: start }, { transform: "translate(0, 0)" }];
    } else if (effect === "pull") frames = [];
    else if (effect === "split") {
      const center = transition.orientation === "horz" ? "inset(50% 0)" : "inset(0 50%)";
      frames = transition.direction === "in" ? [] : [{ clipPath: center }, { clipPath: "inset(0)" }];
    } else if (effect === "doors" || effect === "window") frames = [{ clipPath: effect === "window" ? "inset(50%)" : transition.orientation === "vert" ? "inset(0 50%)" : "inset(50% 0)" }, { clipPath: "inset(0)" }];
    else if (effect === "honeycomb") {
      const { width, height } = incoming.getBoundingClientRect();
      frames = [{ clipPath: honeycombClipPath(width, height, true) }, { clipPath: honeycombClipPath(width, height, false) }];
    } else if (effect === "strips") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: stripsClipPath(width, height, direction, progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "dissolve") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: dissolveClipPath(width, height, progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "comb") {
      const { width, height } = incoming.getBoundingClientRect();
      const orientation = transition.orientation || "horz";
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: combClipPath(width, height, orientation, progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "blinds" || effect === "checker") {
      const { width, height } = incoming.getBoundingClientRect();
      const orientation = transition.orientation || "horz";
      const progressValues = effect === "checker" ? [
        0,
        0.25,
        0.5,
        0.75,
        1
      ] : [
        0,
        0.5,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: gridTransitionClipPath(width, height, orientation, progress, effect === "checker"),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "randomBar") {
      const { width, height } = incoming.getBoundingClientRect();
      const orientation = transition.barOrientation || "horizontal";
      const seed = activeSlide.id;
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: randomBarClipPath(width, height, orientation, progress, seed),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "wheel" || effect === "wheelReverse") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: wheelClipPath(width, height, transition.spokes || 4, progress, effect === "wheelReverse"),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "plus" || effect === "wedge") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: effect === "plus" ? plusTransitionClipPath(width, height, progress) : wedgeTransitionClipPath(width, height, progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "ripple") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.2,
        0.4,
        0.6,
        0.8,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: rippleTransitionClipPath(width, height, transition.direction, progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "glitter") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      frames = progressValues.map((progress, index) => ({
        clipPath: glitterTransitionClipPath(width, height, direction, transition.pattern || "diamond", progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "shred") {
      const { width, height } = incoming.getBoundingClientRect();
      const progressValues = [
        0,
        0.25,
        0.5,
        0.75,
        1
      ];
      shredOutgoing = transition.direction === "out";
      if (shredOutgoing) (_q = underlay.value) == null ? void 0 : _q.classList.add("is-shredding-out");
      frames = progressValues.map((progress, index) => ({
        clipPath: shredTransitionClipPath(width, height, transition.shredPattern || "strip", shredOutgoing ? 1 - progress : progress),
        offset: index / (progressValues.length - 1)
      }));
    } else if (effect === "flythrough") {
      const baseTransform = incoming.style.transform === "none" ? "" : `${incoming.style.transform} `;
      const outgoing2 = (_r = underlay.value) == null ? void 0 : _r.firstElementChild;
      const outgoingBase = (outgoing2 == null ? void 0 : outgoing2.style.transform) && outgoing2.style.transform !== "none" ? `${outgoing2.style.transform} ` : "";
      const incomingTransform = (scale) => `${baseTransform}scale(${scale})`.trim();
      const outgoingTransformAt = (scale) => `${outgoingBase}scale(${scale})`.trim();
      (_s = underlay.value) == null ? void 0 : _s.classList.add("is-flythrough");
      frames = transition.hasBounce ? [
        {
          transform: incomingTransform(direction === "in" ? 2.4 : 0.35),
          opacity: 0
        },
        {
          transform: incomingTransform(1.08),
          opacity: 1,
          offset: 0.82
        },
        {
          transform: incomingTransform(1),
          opacity: 1
        }
      ] : [{
        transform: incomingTransform(direction === "in" ? 2.4 : 0.35),
        opacity: 0
      }, {
        transform: incomingTransform(1),
        opacity: 1
      }];
      outgoingFrames = direction === "in" ? [{
        transform: outgoingTransformAt(1),
        opacity: 1
      }, {
        transform: outgoingTransformAt(0.35),
        opacity: 0
      }] : [{
        transform: outgoingTransformAt(1),
        opacity: 1
      }, {
        transform: outgoingTransformAt(2.4),
        opacity: 0
      }];
    } else if (effect === "warp") {
      const incomingBase = incoming.style.transform && incoming.style.transform !== "none" ? `${incoming.style.transform} ` : "";
      const outgoing2 = (_t = underlay.value) == null ? void 0 : _t.firstElementChild;
      const outgoingBase = (outgoing2 == null ? void 0 : outgoing2.style.transform) && outgoing2.style.transform !== "none" ? `${outgoing2.style.transform} ` : "";
      const warped = (base, x, y, scale) => `${base}perspective(1200px) rotateX(${x}deg) rotateY(${y}deg) skewX(${y / 3}deg) scale(${scale})`.trim();
      if (direction === "out") {
        (_u = underlay.value) == null ? void 0 : _u.classList.add("is-warping-out");
        frames = [{
          transform: incomingBase.trim() || "none",
          opacity: 1
        }, {
          transform: incomingBase.trim() || "none",
          opacity: 1
        }];
        outgoingFrames = [{
          transform: outgoingBase.trim() || "none",
          opacity: 1
        }, {
          transform: warped(outgoingBase, 18, -24, 0.72),
          opacity: 0,
          offset: 1
        }];
      } else frames = [
        {
          transform: warped(incomingBase, -18, 24, 0.72),
          opacity: 0
        },
        {
          transform: warped(incomingBase, 5, -7, 1.04),
          opacity: 1,
          offset: 0.78
        },
        {
          transform: incomingBase.trim() || "none",
          opacity: 1
        }
      ];
    } else if (effect === "vortex") {
      const outgoing2 = (_v = underlay.value) == null ? void 0 : _v.firstElementChild;
      const outgoingBase = (outgoing2 == null ? void 0 : outgoing2.style.transform) && outgoing2.style.transform !== "none" ? `${outgoing2.style.transform} ` : "";
      const rotations = {
        l: {
          axis: "0, 0, 1",
          angle: -540
        },
        r: {
          axis: "0, 0, 1",
          angle: 540
        },
        u: {
          axis: "1, 0, 0",
          angle: -540
        },
        d: {
          axis: "1, 0, 0",
          angle: 540
        },
        ld: {
          axis: "1, 1, 0",
          angle: -540
        },
        lu: {
          axis: "1, -1, 0",
          angle: 540
        },
        rd: {
          axis: "1, -1, 0",
          angle: 540
        },
        ru: {
          axis: "1, 1, 0",
          angle: -540
        }
      };
      const rotation = rotations[direction] || rotations.l;
      const endTransform = `${outgoingBase}perspective(1000px) rotate3d(${rotation.axis}, ${rotation.angle}deg) scale(0.08)`.trim();
      (_w = underlay.value) == null ? void 0 : _w.classList.add("is-vortexing");
      frames = [{
        transform: incoming.style.transform || "none",
        opacity: 1
      }, {
        transform: incoming.style.transform || "none",
        opacity: 1
      }];
      outgoingFrames = [{
        transform: outgoingBase.trim() || "none",
        opacity: 1
      }, {
        transform: endTransform,
        opacity: 0
      }];
    } else if (effect === "circle") frames = [{ clipPath: "circle(0% at 50% 50%)" }, { clipPath: "circle(100% at 50% 50%)" }];
    else if (effect === "diamond") frames = [{ clipPath: "polygon(50% 50%, 50% 50%, 50% 50%, 50% 50%)" }, { clipPath: "polygon(50% -50%, 150% 50%, 50% 150%, -50% 50%)" }];
    else if (effect === "newsflash") frames = [{
      transform: "scale(.2) rotate(360deg)",
      opacity: 0
    }, {
      transform: "scale(1) rotate(0deg)",
      opacity: 1
    }];
    else if (effect === "zoom") frames = [{
      transform: direction === "out" ? "scale(1.5)" : "scale(.5)",
      opacity: 0
    }, {
      transform: "scale(1)",
      opacity: 1
    }];
    else frames = [{ clipPath: direction === "r" ? "inset(0 0 0 100%)" : direction === "u" ? "inset(100% 0 0 0)" : direction === "d" ? "inset(0 0 100% 0)" : "inset(0 100% 0 0)" }, { clipPath: "inset(0 0 0 0)" }];
    const timing = {
      duration: transition.durationMs,
      easing: effect === "zoom" || effect === "flythrough" || effect === "warp" || effect === "vortex" ? "ease-in-out" : "ease-out",
      fill: "both",
      delay: animationDelay
    };
    const outgoing = (_x = underlay.value) == null ? void 0 : _x.firstElementChild;
    let outgoingAnimation;
    if (outgoingFrames && outgoing) outgoingAnimation = outgoing.animate(outgoingFrames, timing);
    if (effect === "fade" && transition.throughBlack) {
      (_y = shell.value) == null ? void 0 : _y.classList.add("is-fading-through-black");
      outgoing == null ? void 0 : outgoing.animate([
        {
          opacity: 1,
          offset: 0
        },
        {
          opacity: 0,
          offset: 0.5
        },
        {
          opacity: 0,
          offset: 1
        }
      ], {
        ...timing,
        easing: "linear"
      });
    }
    if (revealThroughBlack) (_z = shell.value) == null ? void 0 : _z.classList.add("is-fading-through-black");
    let animation;
    if (shredOutgoing && outgoing) animation = outgoing.animate(frames, timing);
    else if (shredOutgoing) {
      (_A = underlay.value) == null ? void 0 : _A.classList.remove("is-shredding-out");
      (_B = underlay.value) == null ? void 0 : _B.replaceChildren();
      return;
    } else if (effect === "pull" && outgoing) {
      const exits = {
        l: "translateX(-100%)",
        r: "translateX(100%)",
        u: "translateY(-100%)",
        d: "translateY(100%)",
        ld: "translate(-100%, 100%)",
        lu: "translate(-100%, -100%)",
        rd: "translate(100%, 100%)",
        ru: "translate(100%, -100%)"
      };
      const baseTransform = getComputedStyle(outgoing).transform;
      const start = baseTransform === "none" ? "translate(0, 0)" : baseTransform;
      animation = outgoing.animate([{ transform: start }, { transform: `${start} ${exits[direction] || exits.l}` }], timing);
    } else if (effect === "pull") {
      (_C = underlay.value) == null ? void 0 : _C.replaceChildren();
      return;
    } else if (effect === "split" && transition.direction === "in" && outgoing) {
      const center = transition.orientation === "horz" ? "inset(50% 0)" : "inset(0 50%)";
      animation = outgoing.animate([{ clipPath: "inset(0)" }, { clipPath: center }], timing);
    } else if (effect === "split" && transition.direction === "in") {
      (_D = underlay.value) == null ? void 0 : _D.replaceChildren();
      return;
    } else {
      if (outgoingTransform && outgoing) {
        const baseTransform = getComputedStyle(outgoing).transform;
        const start = baseTransform === "none" ? "translate(0, 0)" : baseTransform;
        const end = `${start} ${outgoingTransform}`;
        const keyframes = revealThroughBlack ? [
          {
            transform: start,
            opacity: 1
          },
          {
            transform: end,
            opacity: 0,
            offset: 0.5
          },
          {
            transform: end,
            opacity: 0
          }
        ] : [{ transform: start }, { transform: end }];
        outgoing.animate(keyframes, timing);
      }
      animation = incoming.animate(frames, (effect === "fade" || revealThroughBlack) && transition.throughBlack ? {
        ...timing,
        easing: "linear"
      } : timing);
    }
    animation.onfinish = () => {
      var _a2, _b2, _c2, _d2;
      animation.cancel();
      outgoingAnimation == null ? void 0 : outgoingAnimation.cancel();
      if (((_a2 = slide.value) == null ? void 0 : _a2.id) === activeSlide.id) {
        (_b2 = underlay.value) == null ? void 0 : _b2.replaceChildren();
        (_c2 = underlay.value) == null ? void 0 : _c2.classList.remove("is-shredding-out", "is-flythrough", "is-warping-out", "is-vortexing");
        (_d2 = shell.value) == null ? void 0 : _d2.classList.remove("is-fading-through-black");
      }
    };
  }, { flush: "post" });
}

export { SlideStage_default as S, cssFontFamily as c, paragraphEntries as p, scriptSegments as s, useSlideTransition as u };
//# sourceMappingURL=useSlideTransition-DQlUkfdc.mjs.map
