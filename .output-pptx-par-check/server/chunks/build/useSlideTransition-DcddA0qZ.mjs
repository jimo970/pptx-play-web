import { _ as _plugin_vue_export_helper_default } from '../virtual/entry.mjs';
import { watch, nextTick, defineComponent, useId, ref, computed, reactive, mergeProps, unref, createVNode, resolveDynamicComponent, withCtx, openBlock, createBlock, Fragment, renderList, createCommentVNode, toDisplayString, createTextVNode, useSSRContext } from 'vue';
import { ssrRenderAttrs, ssrRenderList, ssrRenderAttr, ssrRenderClass, ssrRenderStyle, ssrInterpolate, ssrRenderVNode } from 'vue/server-renderer';

//#region app/utils/chart-number-format.ts
var numberFormatColors = {
	black: "#000000",
	blue: "#0000ff",
	cyan: "#00ffff",
	green: "#008000",
	magenta: "#ff00ff",
	red: "#ff0000",
	white: "#ffffff",
	yellow: "#ffff00"
};
var defaultIndexedColors = [
	"000000",
	"FFFFFF",
	"FF0000",
	"00FF00",
	"0000FF",
	"FFFF00",
	"FF00FF",
	"00FFFF",
	"800000",
	"008000",
	"000080",
	"808000",
	"800080",
	"008080",
	"C0C0C0",
	"808080",
	"9999FF",
	"993366",
	"FFFFCC",
	"CCFFFF",
	"660066",
	"FF8080",
	"0066CC",
	"CCCCFF",
	"000080",
	"FF00FF",
	"FFFF00",
	"00FFFF",
	"800080",
	"800000",
	"008080",
	"0000FF",
	"00CCFF",
	"CCFFFF",
	"CCFFCC",
	"FFFF99",
	"99CCFF",
	"FF99CC",
	"CC99FF",
	"FFCC99",
	"3366FF",
	"33CCCC",
	"99CC00",
	"FFCC00",
	"FF9900",
	"FF6600",
	"666699",
	"969696",
	"003366",
	"339966",
	"003300",
	"333300",
	"993300",
	"993366",
	"333399",
	"333333"
];
function indexedColor(index, indexedColors) {
	const hex = (indexedColors?.[index + 7])?.trim().replace(/^#/, "").slice(-6);
	return hex && /^[\da-f]{6}$/i.test(hex) ? `#${hex.toLowerCase()}` : `#${defaultIndexedColors[index - 1].toLowerCase()}`;
}
function numberFormatColor(code, indexedColors) {
	if (!code) return void 0;
	const match = /^\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color(?:[1-9]|[1-4]\d|5[0-6]))\]/i.exec(code);
	if (!match) return void 0;
	const colorIndex = /^Color(\d+)$/i.exec(match[1]);
	return colorIndex ? indexedColor(Number(colorIndex[1]), indexedColors) : numberFormatColors[match[1].toLowerCase()];
}
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
	const nextTick = majorTimeUnit === "days" ? minimum + majorUnit * 101 : nextCalendarDate ? serialFromCalendarDate(nextCalendarDate, date1904) : void 0;
	return nextTick !== void 0 && Number.isFinite(nextTick) && nextTick <= maximum + 1e-9 ? void 0 : ticks;
}
function parseDateFormat(formatCode) {
	if (!formatCode || formatCode.length > 128) return void 0;
	const tokens = [];
	for (let index = 0; index < formatCode.length;) {
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
		if (char === "\"") {
			const end = formatCode.indexOf("\"", index + 1);
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
			if (!/^\$(?:[^\]]*)-[0-9a-f]+$/i.test(bracket) && !/^[a-z]+$/i.test(bracket) && !/^Color(?:[1-9]|[1-4]\d|5[0-6])$/i.test(bracket)) return void 0;
			index = end + 1;
			continue;
		}
		if (/[dmyhs]/i.test(char)) {
			const code = char.toLowerCase();
			let end = index + 1;
			while (formatCode[end]?.toLowerCase() === code) end++;
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
			if (previousToken?.kind === "date" && previousToken.code === "s") {
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
		return hasClock && (previous?.kind === "date" && previous.code === "h" || next?.kind === "date" && next.code === "s") ? {
			...token,
			length: -token.length
		} : token;
	});
}
function formatDateNumber(value, tokens, date1904) {
	if (!Number.isFinite(value) || value < 0 || value > 2958465) return void 0;
	const fractionDigits = tokens.find((token) => token.kind === "fraction")?.digits;
	const precisionDigits = Math.max(fractionDigits ?? 3, 3);
	const unitsPerSecond = 10 ** precisionDigits;
	const unitsPerDay = 86400 * unitsPerSecond;
	const displayStep = 10 ** (precisionDigits - (fractionDigits ?? 3));
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
			const value = ampm ? hour % 12 || 12 : hour;
			return token.length === 1 ? String(value) : String(value).padStart(2, "0");
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
		if (char === "\"") quoted = !quoted;
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
		} else if (char === "\"") quoted = !quoted;
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
	if (!code) return null;
	const colorMatch = /^\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color(?:[1-9]|[1-4]\d|5[0-6]))\]/i.exec(code);
	const colorIndex = colorMatch ? /^Color(\d+)$/i.exec(colorMatch[1]) : void 0;
	const color = colorMatch && !colorIndex ? numberFormatColors[colorMatch[1].toLowerCase()] : void 0;
	if (colorMatch) code = code.slice(colorMatch[0].length);
	const conditionMatch = /^\[(<=|>=|<>|=|<|>)(-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)\]/.exec(code);
	const conditionValue = conditionMatch ? Number(conditionMatch[2]) : void 0;
	if (conditionMatch && !Number.isFinite(conditionValue)) return void 0;
	const condition = conditionMatch ? {
		operator: conditionMatch[1],
		value: conditionValue
	} : void 0;
	if (conditionMatch) code = code.slice(conditionMatch[0].length);
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
		if (char === "\"") quoted = !quoted;
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
		const suffixStart = denominatorEnd + (percentMatch?.[0].length || 0);
		const prefix = parseNumberFormatLiteral(code.slice(0, start));
		const percentSpacing = percentMatch ? parseNumberFormatLiteral(percentMatch[1]) : "";
		const remainingSuffix = parseNumberFormatLiteral(code.slice(suffixStart));
		if (prefix === void 0 || percentSpacing === void 0 || remainingSuffix === void 0 || !percent && rawSuffix.includes("%") || numerator.length > 3 || fixedDenominator === void 0 && denominator.length > 3 || whole && whole.replaceAll(",", "").length > 15 || fixedDenominator !== void 0 && (!Number.isSafeInteger(fixedDenominator) || fixedDenominator < 1 || fixedDenominator > 999999)) return void 0;
		const suffix = `${percentSpacing}${percent ? "%" : ""}${remainingSuffix}`;
		return {
			options: {
				useGrouping: Boolean(whole?.includes(",")),
				minimumIntegerDigits: 1,
				minimumFractionDigits: 0,
				maximumFractionDigits: 0
			},
			parentheses,
			prefix,
			suffix,
			scale: 0,
			...color ? { color } : {},
			...colorIndex ? { colorIndex: Number(colorIndex[1]) } : {},
			...condition ? { condition } : {},
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
	if ((match[2]?.length || 0) > 100) return void 0;
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
		...color ? { color } : {},
		...colorIndex ? { colorIndex: Number(colorIndex[1]) } : {},
		...condition ? { condition } : {},
		...scientific ? { scientific } : {}
	};
}
function conditionsCoverAllValues(left, right) {
	if (left.value !== right.value) return false;
	const operators = [left.operator, right.operator];
	return operators.includes("<=") && operators.includes(">") || operators.includes("<") && operators.includes(">=") || operators.includes("<>") && operators.includes("=");
}
function parseNumberFormat(formatCode) {
	const codes = splitNumberFormatSections(formatCode);
	if (!codes || codes.length > 3 || !codes[0]) return void 0;
	const sections = codes.map(parseNumberFormatSection);
	if (sections.some((section) => section === void 0)) return void 0;
	const parsedSections = sections;
	const conditionalIndexes = parsedSections.flatMap((section, index) => section?.condition ? [index] : []);
	if (!conditionalIndexes.length) return parsedSections;
	return (codes.length === 2 ? conditionalIndexes[0] === 0 && (conditionalIndexes.length === 1 || conditionalIndexes[1] === 1 && parsedSections[0]?.condition !== void 0 && parsedSections[1]?.condition !== void 0 && conditionsCoverAllValues(parsedSections[0].condition, parsedSections[1].condition)) : codes.length === 3 ? conditionalIndexes.length === 2 && conditionalIndexes[0] === 0 && conditionalIndexes[1] === 1 : false) && !parsedSections.some((section) => section?.parentheses) ? parsedSections : void 0;
}
function matchesNumberCondition(value, condition) {
	switch (condition.operator) {
		case "<": return value < condition.value;
		case "<=": return value <= condition.value;
		case "=": return value === condition.value;
		case ">": return value > condition.value;
		case ">=": return value >= condition.value;
		case "<>": return value !== condition.value;
	}
}
function formatScientificNumber(value, section) {
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
	return `${formatted}${`${scientific.symbol}${exponentSign}${String(Math.abs(exponent)).padStart(scientific.minimumDigits, "0")}`}${section.options.style === "percent" ? new Intl.NumberFormat(void 0, {
		style: "percent",
		maximumFractionDigits: 0
	}).formatToParts(.01).find((part) => part.type === "percentSign")?.value || "%" : ""}`;
}
function formatChartNumberWithColor(value, formatCode, date1904 = false, indexedColors) {
	if (!formatCode || formatCode.toLowerCase() === "general") return { text: String(value) };
	const sections = splitNumberFormatSections(formatCode);
	const dateSection = sections && (value < 0 ? sections[1] : value === 0 ? sections[2] || sections[0] : sections[0]);
	const dateTokens = dateSection ? parseDateFormat(dateSection) : void 0;
	if (dateTokens) {
		const text = formatDateNumber(Math.abs(value), dateTokens, date1904) ?? String(value);
		const color = numberFormatColor(dateSection, indexedColors);
		return {
			text,
			...color ? { color } : {}
		};
	}
	const numberSections = parseNumberFormat(formatCode);
	if (!numberSections) return { text: String(value) };
	let section = numberSections[0];
	let formattedValue = value;
	if (numberSections.some((candidate) => candidate?.condition)) {
		const matching = numberSections.slice(0, 2).find((candidate) => candidate?.condition && matchesNumberCondition(value, candidate.condition));
		const fallback = numberSections.find((candidate) => candidate && !candidate.condition);
		const selected = matching || fallback;
		if (!selected) return { text: String(value) };
		section = selected;
	} else if (value < 0 && numberSections.length > 1) {
		const negative = numberSections[1];
		if (!negative) return { text: "" };
		section = negative;
		formattedValue = Math.abs(value);
	} else if (value === 0 && numberSections.length > 2) {
		const zero = numberSections[2];
		if (!zero) return { text: "" };
		section = zero;
	}
	let scaledValue = formattedValue;
	for (let index = 0; index < section.scale; index++) scaledValue /= 1e3;
	const formatted = section.fraction ? formatFractionNumber(scaledValue, section.fraction) : section.scientific ? formatScientificNumber(scaledValue, section) : new Intl.NumberFormat(void 0, section.options).format(scaledValue);
	if (formatted === void 0) return { text: String(value) };
	const valueText = `${section.prefix}${formatted}${section.suffix}`;
	const text = section.parentheses ? `(${valueText})` : valueText;
	const color = section.color ?? (section.colorIndex ? indexedColor(section.colorIndex, indexedColors) : void 0);
	return {
		text,
		...color ? { color } : {}
	};
}
//#endregion
//#region app/utils/paragraph-bullets.ts
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
	const digits = "零一二三四五六七八九";
	const units = traditional ? [
		"",
		"萬",
		"億",
		"兆",
		"京"
	] : [
		"",
		"万",
		"亿",
		"兆",
		"京"
	];
	const groupUnits = [
		"",
		"十",
		"百",
		"千"
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
	if (type === "arabicDbPlain") return scriptDigits(value, "０１２３４５６７８９");
	let label = String(value);
	if (type.startsWith("ea1Chs")) label = chineseNumber(value, false);
	else if (type.startsWith("ea1Cht")) label = chineseNumber(value, true);
	else if (type.startsWith("thaiAlpha")) label = alphabeticNumber(value, "กขคฆงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ");
	else if (type.startsWith("thaiNum")) label = scriptDigits(value, "๐๑๒๓๔๕๖๗๘๙");
	else if (type.startsWith("hindiNum")) label = scriptDigits(value, "०१२३४५६७८९");
	else if (type === "arabicDbPeriod") return `${scriptDigits(value, "０１２３４５６７８９")}．`;
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
		const bullet = paragraph.bullet;
		if (!bullet || bullet.kind !== "number") {
			counters.clear();
			return {
				paragraph,
				marker: bullet?.kind === "char" ? bullet.char : ""
			};
		}
		const key = `${bullet.type}:${paragraph.level || 0}`;
		const number = counters.get(key) ?? bullet.startAt ?? 1;
		counters.set(key, number + 1);
		return {
			paragraph,
			marker: numberedMarker(number, bullet.type)
		};
	});
}
//#endregion
//#region app/utils/color.ts
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
//#endregion
//#region app/utils/text-fonts.ts
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
		if (last?.script === script) last.text += character;
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
//#endregion
//#region app/utils/pptx-render-utils.ts
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
var SVG_POLYGON_GEOMETRIES = /* @__PURE__ */ new Set([
	"leftBracket",
	"rightBracket",
	"bracketPair",
	"leftBrace",
	"rightBrace",
	"bracePair",
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
	"pieWedge",
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
]);
function isSvgPolygonGeometry(geometry) {
	return SVG_POLYGON_GEOMETRIES.has(geometry);
}
//#endregion
//#region app/utils/animation-timing.ts
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
//#endregion
//#region app/utils/transitionMasks.ts
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
			amount = Math.max(0, Math.min(1, (progress - step / Math.max(1, count - 1) * .72) / .28));
		}
		const left = column * cellWidth - .5;
		const right = (column + 1) * cellWidth + .5;
		const top = row * cellHeight - .5;
		const bottom = (row + 1) * cellHeight + .5;
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
	const window = .3;
	return `path("${order.map((rank, band) => {
		const start = rank / Math.max(1, count - 1) * .7;
		const reveal = Math.max(0, Math.min(1, (amount - start) / window));
		const center = (band + .5) * bandSize;
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
		const crossStart = index * barSize - .5;
		const crossEnd = (index + 1) * barSize + .5;
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
	const staggerWindow = .38;
	return `path("${Array.from({ length: count }, (_, band) => {
		const start = (reverseRows ? 7 - band : band) / 7 * .62;
		const reveal = Math.max(0, Math.min(1, (amount - start) / staggerWindow));
		const top = band * bandHeight - .5;
		const bottom = (band + 1) * bandHeight + .5;
		const base = reverseColumns ? width : 0;
		const travel = width * reveal;
		const slant = reveal > 0 && reveal < 1 ? bandHeight * .75 : 0;
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
	const amplitude = Math.min(width, height) * .025 * Math.sin(Math.PI * amount);
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
		const local = Math.max(0, Math.min(1, (amount - rank * .72) / .28));
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
		const local = Math.max(0, Math.min(1, (amount - rank * .68) / .32));
		const overlap = local ? .25 : 0;
		const halfWidth = cellWidth * local / 2 + overlap;
		const halfHeight = cellHeight * local / 2 + overlap;
		const centerX = (column + .5) * cellWidth;
		const centerY = (row + .5) * cellHeight;
		paths.push(`M${(centerX - halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} L${(centerX - halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} Z`);
	}
	return `path("${paths.join(" ")}")`;
}
function plusPath(width, height, amount) {
	const cx = width / 2, cy = height / 2;
	const reachX = width * amount / 2;
	const reachY = height * amount / 2;
	const arm = amount * .28 + Math.pow(amount, 12) * .72;
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
		const left = column * cellWidth - .5;
		const right = (column + 1) * cellWidth + .5;
		const top = row * cellHeight - .5;
		const bottom = (row + 1) * cellHeight + .5;
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
		const centerAngle = -Math.PI / 2 + (spoke + .5) * Math.PI * 2 / spokes;
		const halfAngle = Math.PI / spokes * progress;
		const points = Array.from({ length: 13 }, (_, point) => {
			const angle = reverse ? -Math.PI / 2 + (spoke + 1) * Math.PI * 2 / spokes - Math.PI * 2 / spokes * progress * point / pointsPerSpoke : centerAngle - halfAngle + halfAngle * 2 * point / pointsPerSpoke;
			return `L${(centerX + Math.cos(angle) * radius).toFixed(2)} ${(centerY + Math.sin(angle) * radius).toFixed(2)}`;
		});
		return `M${centerX.toFixed(2)} ${centerY.toFixed(2)} ${points.join(" ")} Z`;
	}).join(" ")}")`;
}
//#endregion
//#region app/components/SlideStage.vue?vue&type=script&setup=true&lang.ts
var EMU_PER_POINT = 12700;
var SlideStage_vue_vue_type_script_setup_true_lang_default = /*@__PURE__*/ defineComponent({
	__name: "SlideStage",
	__ssrInlineRender: true,
	props: {
		slide: {},
		width: {},
		height: {},
		number: {},
		animationStep: { default: 0 },
		showAnimations: {
			type: Boolean,
			default: true
		},
		showNarration: {
			type: Boolean,
			default: true
		},
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
		const svgStrokeDashPatterns = {
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
		function svgStrokeDashPattern(dash, cap) {
			const pattern = Array.isArray(dash) ? dash : svgStrokeDashPatterns[dash || ""];
			if (!pattern) return void 0;
			const capLength = cap === "round" || !Array.isArray(dash) && cap === "square" ? 1 : 0;
			return pattern.map((length, index) => {
				const extension = Math.min(capLength, pattern[index % 2 ? index - 1 : index]);
				return length + (index % 2 ? extension : -extension);
			});
		}
		function defaultLegendOrigin(position, axis) {
			if (axis === "x") return position === "left" ? 88 : position === "right" ? 780 : position === "topRight" ? 760 : 88;
			return position === "left" || position === "right" ? 105 : position === "top" ? 77 : position === "topRight" ? 92 : 530;
		}
		function chartMarkerPath(marker) {
			const { x, y } = marker;
			const radius = marker.radius ?? 5;
			const symbol = marker.symbol ?? "circle";
			if (symbol === "square") return `M ${x - radius} ${y - radius} h ${radius * 2} v ${radius * 2} h ${-radius * 2} Z`;
			if (symbol === "diamond") return `M ${x} ${y - radius} L ${x + radius} ${y} L ${x} ${y + radius} L ${x - radius} ${y} Z`;
			if (symbol === "triangle") return `M ${x} ${y - radius} L ${x + radius} ${y + radius} L ${x - radius} ${y + radius} Z`;
			if (symbol === "dash") return `M ${x - radius} ${y - radius * .22} h ${radius * 2} v ${radius * .44} h ${-radius * 2} Z`;
			if (symbol === "plus") {
				const arm = radius * .28;
				return `M ${x - arm} ${y - radius} H ${x + arm} V ${y - arm} H ${x + radius} V ${y + arm} H ${x + arm} V ${y + radius} H ${x - arm} V ${y + arm} H ${x - radius} V ${y - arm} H ${x - arm} Z`;
			}
			if (symbol === "x") {
				const arm = radius * .28;
				return `M ${x - radius} ${y - radius + arm} L ${x - radius + arm} ${y - radius} L ${x} ${y - arm} L ${x + radius - arm} ${y - radius} L ${x + radius} ${y - radius + arm} L ${x + arm} ${y} L ${x + radius} ${y + radius - arm} L ${x + radius - arm} ${y + radius} L ${x} ${y + arm} L ${x - radius + arm} ${y + radius} L ${x - radius} ${y + radius - arm} L ${x - arm} ${y} Z`;
			}
			if (symbol === "star") return Array.from({ length: 10 }, (_, index) => {
				const angle = -Math.PI / 2 + index * Math.PI / 5;
				const pointRadius = radius * (index % 2 ? .45 : 1);
				return `${index ? "L" : "M"} ${x + Math.cos(angle) * pointRadius} ${y + Math.sin(angle) * pointRadius}`;
			}).join(" ") + " Z";
			return "";
		}
		function smoothChartLinePath(points, smooth) {
			const first = points[0];
			if (!first) return "";
			let path = `M ${first.x} ${first.y}`;
			for (let index = 0; index < points.length - 1; index++) {
				const start = points[index];
				const end = points[index + 1];
				if (!start || !end) continue;
				if (!smooth || points.length < 3) {
					path += ` L ${end.x} ${end.y}`;
					continue;
				}
				const previous = points[Math.max(0, index - 1)] ?? start;
				const next = points[Math.min(index + 2, points.length - 1)] ?? end;
				path += ` C ${start.x + (end.x - previous.x) / 6} ${start.y + (end.y - previous.y) / 6} ${end.x - (next.x - start.x) / 6} ${end.y - (next.y - start.y) / 6} ${end.x} ${end.y}`;
			}
			return path;
		}
		function makeChartVisual(chart, elementId, chartFrameWidthEmu) {
			const chartStrokeWidth = (widthEmu) => widthEmu === void 0 ? void 0 : (widthEmu || 9525) * 1e3 / Math.max(chartFrameWidthEmu, 1);
			const chartStrokeStyle = (series, index, width) => {
				const dash = index === void 0 ? series.strokeDash : series.pointStrokeDashes?.[index] ?? series.strokeDash;
				const cap = index === void 0 ? series.strokeCap : series.pointStrokeCaps?.[index] ?? series.strokeCap;
				const pattern = svgStrokeDashPattern(dash, cap);
				return {
					...pattern ? { strokeDasharray: pattern.map((length) => length * width).join(" ") } : {},
					...cap ? { strokeLinecap: cap } : {}
				};
			};
			const combo = chart.series.some((series) => series.chartType);
			const seriesType = (series) => series.chartType || chart.type;
			const axisEntryForSeries = (series) => chart.valueAxes?.find((item) => item.id === series.valueAxisId);
			const primaryAxisId = chart.valueAxes?.find((item) => item.primary)?.id;
			const axisValuesForId = (axisId) => chart.series.filter((series) => !axisId || (series.valueAxisId || primaryAxisId) === axisId).flatMap((series) => series.values.filter((value) => value !== null && value !== void 0));
			const scatter = chart.type === "scatter";
			const bubble = chart.type === "bubble";
			const radar = chart.type === "radar";
			const stock = chart.type === "stock" ? chart.stock : void 0;
			const xyChart = scatter || bubble;
			const position = chart.legend?.position;
			const plotPosition = chart.legend?.overlay ? void 0 : position;
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
			const radarRadiusX = plot.width * .38;
			const radarRadiusY = plot.height * .38;
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
			].filter((index) => index !== void 0).flatMap((index) => chart.series[index]?.values.filter((value) => value !== null && value !== void 0) || []) : combo ? axisValuesForId(primaryAxisId) : chart.series.flatMap((series) => series.values.filter((value) => value !== null && value !== void 0));
			const stacking = !combo && [
				"bar",
				"line",
				"area"
			].includes(chart.type) && chart.grouping && chart.grouping !== "clustered" ? getPptxChartStacking(chart.series, chart.grouping) : void 0;
			const stackedSegments = /* @__PURE__ */ new Map();
			if (combo) {
				const groupIndexes = /* @__PURE__ */ new Map();
				chart.series.forEach((series, index) => {
					const groupIndex = series.chartGroupIndex ?? 0;
					groupIndexes.set(groupIndex, [...groupIndexes.get(groupIndex) || [], index]);
				});
				for (const [groupIndex, indexes] of groupIndexes) {
					const grouping = chart.series[indexes[0]].grouping || "clustered";
					if (grouping === "clustered") continue;
					const groupStacking = getPptxChartStacking(indexes.map((index) => chart.series[index]), grouping);
					indexes.forEach((index, seriesIndex) => stackedSegments.set(index, groupStacking.segments[seriesIndex]));
				}
			}
			const stackedSegmentFor = (seriesIndex) => combo ? stackedSegments.get(seriesIndex) : stacking?.segments[seriesIndex];
			const axisScaleFraction = (value, minimum, maximum, axis) => {
				if (!axis?.logBase) return (value - minimum) / (maximum - minimum);
				return (Math.log(value > 0 ? value : minimum) - Math.log(minimum)) / (Math.log(maximum) - Math.log(minimum));
			};
			const axisMajorTickValues = (axis, minimum, maximum) => {
				if (!axis?.logBase) {
					const unit = axis?.majorUnit;
					return unit ? Array.from({ length: Math.min(101, Math.floor((maximum - minimum) / unit + 1e-9) + 1) }, (_, index) => Number((minimum + index * unit).toPrecision(12))) : Array.from({ length: 5 }, (_, index) => minimum + (maximum - minimum) * index / 4);
				}
				const logarithmBase = Math.log(axis.logBase);
				const minimumExponent = Math.log(minimum) / logarithmBase;
				const maximumExponent = Math.log(maximum) / logarithmBase;
				const step = axis.majorUnit ?? 1;
				const firstExponent = Math.ceil((minimumExponent - 1e-9) / step) * step;
				const tickCount = Math.min(101, Math.floor((maximumExponent - firstExponent) / step + 1e-9) + 1);
				const values = Array.from({ length: Math.max(0, tickCount) }, (_, index) => axis.logBase ** (firstExponent + index * step)).filter((value) => Number.isFinite(value) && value >= minimum * (1 - 1e-9) && value <= maximum * (1 + 1e-9));
				return values.length ? values : [minimum, maximum];
			};
			const axisBoundsForSeries = (series) => {
				const axis = axisEntryForSeries(series)?.axis || chart.valueAxis;
				if (!combo && stacking) {
					const minimum = chart.valueAxis?.min ?? stacking.min;
					return {
						axis,
						minimum,
						upper: chart.valueAxis?.max ?? (stacking.max === minimum ? minimum + 1 : stacking.max)
					};
				}
				const values = axisValuesForId(series.valueAxisId || primaryAxisId);
				const dataMin = values.length ? axis?.logBase ? Math.min(...values) : Math.min(0, ...values) : axis?.logBase ? 1 : 0;
				const dataMax = values.length ? Math.max(...values, dataMin) : axis?.logBase ? axis.logBase : 1;
				const minimum = axis?.min ?? dataMin;
				const fallbackUpper = axis?.logBase ? minimum * axis.logBase : minimum + 1;
				return {
					axis,
					minimum,
					upper: axis?.max ?? (dataMax === minimum ? fallbackUpper : dataMax)
				};
			};
			const seriesValuePosition = (series, value) => {
				const { axis, minimum: axisMinimum, upper: axisUpper } = axisBoundsForSeries(series);
				const fraction = axisScaleFraction(value, axisMinimum, axisUpper, axis);
				const axisFraction = axis?.reverse ? 1 - fraction : fraction;
				return chart.direction === "horizontal" ? plot.left + plot.width * axisFraction : plot.top + plot.height * (1 - axisFraction);
			};
			const dataMin = stacking?.min ?? (xyChart && values.length ? Math.min(...values) : Math.min(0, ...values));
			const dataMax = stacking?.max ?? (xyChart && values.length ? Math.max(...values) : Math.max(0, ...values));
			const minimum = chart.valueAxis?.min ?? dataMin;
			const upper = chart.valueAxis?.max ?? (dataMax === dataMin ? dataMin + 1 : dataMax);
			const xValues = chart.series.flatMap((series) => (series.xValues || []).filter((value) => value !== null && value !== void 0));
			const xDataMin = xyChart && xValues.length ? Math.min(...xValues) : 0;
			const xDataMax = xyChart && xValues.length ? Math.max(...xValues) : 1;
			const xMinimum = chart.xAxis?.min ?? (xDataMin === xDataMax ? xDataMin - (Math.abs(xDataMin) * .05 || .5) : xDataMin);
			const xUpper = chart.xAxis?.max ?? (xDataMin === xDataMax ? xDataMax + (Math.abs(xDataMax) * .05 || .5) : xDataMax);
			const categoryCount = Math.max(chart.categories.length, 1);
			const formatCategoryAxisLabel = (text) => {
				if (!chart.categoryAxisNumberFormat || !text.trim()) return { text };
				const value = Number(text);
				return Number.isFinite(value) ? formatChartNumberWithColor(value, chart.categoryAxisNumberFormat, chart.date1904, chart.indexedColors) : { text };
			};
			const horizontalValueAxis = chart.direction === "horizontal";
			const valueAxisSide = chart.valueAxis?.position ?? (horizontalValueAxis ? "b" : "l");
			const xAxisSide = chart.xAxis?.position ?? "b";
			const categoryAxisSide = chart.categoryAxisPosition ?? (horizontalValueAxis ? "l" : "b");
			const valueTickLabelPosition = chart.valueAxis?.tickLabelPosition ?? "nextTo";
			const categoryTickLabelPosition = chart.categoryAxisTickLabelPosition ?? "nextTo";
			const categoryAxisLabelStep = chart.categoryAxisTickLabelSkip ?? 1;
			const categoryAxisTickStep = chart.categoryAxisTickMarkSkip ?? 1;
			const showCategoryAxisLabel = (index) => index % categoryAxisLabelStep === 0;
			const showCategoryAxisTick = (index) => index % categoryAxisTickStep === 0;
			const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
			const dateValues = chart.dateAxis?.values;
			const dateMinimum = chart.dateAxis?.min ?? (dateValues?.length ? Math.min(...dateValues) : 0);
			const dateMaximum = chart.dateAxis?.max ?? (dateValues?.length ? Math.max(...dateValues) : 1);
			const hasDateAxisScale = dateValues?.length === categoryCount && dateMaximum > dateMinimum;
			const dateAxisMajorTickValues = hasDateAxisScale && chart.dateAxis?.majorUnit !== void 0 && chart.dateAxis.majorTimeUnit ? getChartDateAxisMajorTicks(dateMinimum, dateMaximum, chart.dateAxis.majorUnit, chart.dateAxis.majorTimeUnit, chart.date1904) : void 0;
			const dateAxisMinorTickValues = hasDateAxisScale && chart.dateAxis?.minorUnit !== void 0 && chart.dateAxis.minorTimeUnit ? getChartDateAxisMajorTicks(dateMinimum, dateMaximum, chart.dateAxis.minorUnit, chart.dateAxis.minorTimeUnit, chart.date1904) : void 0;
			const dateAxisFraction = (value) => {
				const fraction = clamp((value - dateMinimum) / (dateMaximum - dateMinimum), 0, 1);
				return chart.dateAxis?.reverse ? 1 - fraction : fraction;
			};
			const categoryFraction = (index) => {
				if (!hasDateAxisScale || dateValues?.[index] === void 0) return (index + .5) / categoryCount;
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
				const fraction = axisScaleFraction(value, minimum, upper, chart.valueAxis);
				return chart.valueAxis?.reverse ? 1 - fraction : fraction;
			};
			const xAxisFraction = (value) => {
				const fraction = axisScaleFraction(value, xMinimum, xUpper, chart.xAxis);
				return chart.xAxis?.reverse ? 1 - fraction : fraction;
			};
			const xPosition = (value) => plot.left + plot.width * xAxisFraction(value);
			const valuePosition = (value) => plot.top + plot.height * (1 - axisFraction(value));
			const withoutMajorTickValues = (values, minimum, maximum, majorValues) => values.filter((value) => value > minimum + Number.EPSILON * Math.abs(minimum) && value < maximum - Number.EPSILON * Math.abs(maximum) && !majorValues.some((major) => Math.abs(value - major) <= 1e-9 * Math.max(Math.abs(value), Math.abs(major), 1e-300)));
			const minorAxisTickValues = (majorValues, minimum, maximum, minorUnit, logBase) => {
				const values = [];
				if (logBase) {
					const logarithmBase = Math.log(logBase);
					const minimumExponent = Math.log(minimum) / logarithmBase;
					const maximumExponent = Math.log(maximum) / logarithmBase;
					if (minorUnit && Number.isFinite(minorUnit) && minorUnit > 0) {
						const firstExponent = (Math.floor(minimumExponent / minorUnit + 1e-9) + 1) * minorUnit;
						const count = Math.min(512, Math.floor((maximumExponent - firstExponent) / minorUnit + 1e-9) + 1);
						for (let index = 0; index < count; index++) values.push(logBase ** (firstExponent + index * minorUnit));
					} else if (Number.isInteger(logBase) && logBase <= 16) for (let exponent = Math.floor(minimumExponent); exponent < maximumExponent && values.length < 512; exponent++) {
						const decade = logBase ** exponent;
						for (let multiplier = 2; multiplier < logBase && values.length < 512; multiplier++) values.push(multiplier * decade);
					}
					else for (let index = 1; index < majorValues.length && values.length < 512; index++) {
						const start = majorValues[index - 1];
						const end = majorValues[index];
						for (let part = 1; part < 5 && values.length < 512; part++) values.push(Math.exp(Math.log(start) + (Math.log(end) - Math.log(start)) * part / 5));
					}
				} else if (minorUnit && Number.isFinite(minorUnit) && minorUnit > 0) {
					const count = Math.min(512, Math.floor((maximum - minimum) / minorUnit + 1e-9));
					for (let index = 1; index <= count; index++) values.push(Number((minimum + index * minorUnit).toPrecision(12)));
				} else for (let index = 1; index < majorValues.length; index++) {
					const start = majorValues[index - 1];
					const end = majorValues[index];
					for (let part = 1; part < 5; part++) values.push(start + (end - start) * part / 5);
				}
				return withoutMajorTickValues(values, minimum, maximum, majorValues);
			};
			const categoryCrossingFraction = chart.categoryAxisCrossesAt !== void 0 ? (clamp(chart.categoryAxisCrossesAt, 1, categoryCount) - .5) / categoryCount : chart.valueAxis?.crosses === "min" ? 0 : chart.valueAxis?.crosses === "max" ? 1 : void 0;
			const xyValueAxisCrossingX = xPosition(clamp(chart.valueAxis?.crossesAt ?? (chart.valueAxis?.crosses === "min" ? xMinimum : chart.valueAxis?.crosses === "max" ? xUpper : clamp(0, xMinimum, xUpper)), xMinimum, xUpper));
			const xyXAxisCrossingY = valuePosition(clamp(chart.xAxis?.crossesAt ?? (chart.xAxis?.crosses === "min" ? minimum : chart.xAxis?.crosses === "max" ? upper : clamp(0, minimum, upper)), minimum, upper));
			const valueAxisCrossingPosition = categoryCrossingFraction === void 0 ? void 0 : horizontalValueAxis ? plot.top + plot.height * categoryCrossingFraction : plot.left + plot.width * categoryCrossingFraction;
			const categoryAxisCrossingValue = chart.valueAxis?.crossesAt ?? (chart.categoryAxisCrosses === "min" ? minimum : chart.categoryAxisCrosses === "max" ? upper : clamp(0, minimum, upper));
			const categoryAxisCrossingCoordinate = horizontalValueAxis ? plot.left + plot.width * axisFraction(clamp(categoryAxisCrossingValue, minimum, upper)) : valuePosition(clamp(categoryAxisCrossingValue, minimum, upper));
			const valueAxisLineCoordinate = xyChart ? xyValueAxisCrossingX : valueAxisCrossingPosition ?? (horizontalValueAxis ? valueAxisSide === "t" ? plot.top : plot.top + plot.height : valueAxisSide === "r" ? plot.left + plot.width : plot.left);
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
				const x = radar ? radarCenterX : plot.left + plot.width * axisFraction(value);
				const y = radar ? radarCenterY - radarRadiusY * axisFraction(value) : valuePosition(value);
				const formattedLabel = formatChartNumberWithColor(value, chart.valueAxis?.numberFormat ?? (chart.grouping === "percentStacked" ? "0%" : void 0), chart.date1904, chart.indexedColors);
				return {
					value,
					label: formattedLabel.text,
					...formattedLabel.color ? { color: formattedLabel.color } : {},
					showLabel: valueTickLabelPosition !== "none" && !chart.valueAxis?.deleted,
					x,
					y,
					labelX: radar ? radarCenterX - 10 : horizontalValueAxis ? x : valueTickLabelPosition === "high" ? plot.left + plot.width + 10 : valueTickLabelPosition === "low" ? plot.left - 10 : valueAxisLineCoordinate + (valueAxisSide === "r" ? 10 : -10),
					labelY: radar ? y + 5 : horizontalValueAxis ? valueTickLabelPosition === "high" ? plot.top + plot.height + 22 : valueTickLabelPosition === "low" ? plot.top - 10 : valueAxisLineCoordinate + (valueAxisSide === "t" ? -10 : 22) : y + 5,
					labelAnchor: radar ? "end" : horizontalValueAxis ? "middle" : valueTickLabelPosition === "high" ? "start" : valueTickLabelPosition === "low" ? "end" : valueAxisSide === "r" ? "start" : "end"
				};
			});
			const secondaryValueAxes = (chart.valueAxes || []).filter((item) => !item.primary).map(({ id, axis }) => {
				const values = axisValuesForId(id);
				const minimum = axis.min ?? (values.length ? Math.min(0, ...values) : 0);
				const upper = axis.max ?? Math.max(minimum + 1, 0, ...values);
				const horizontal = horizontalValueAxis;
				const side = axis.position ?? (horizontal ? "t" : "r");
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
				const axisTickValues = axisMajorTickValues(axis, minimum, upper);
				const tickPosition = (value) => {
					const fraction = axisScaleFraction(value, minimum, upper, axis);
					const normalized = axis.reverse ? 1 - fraction : fraction;
					return horizontal ? plot.left + plot.width * normalized : plot.top + plot.height * (1 - normalized);
				};
				const labelPosition = axis.tickLabelPosition ?? "nextTo";
				return {
					id,
					horizontal,
					line: axisLine,
					ticks: axisTickValues.map((value) => {
						const position = tickPosition(value);
						const formattedLabel = formatChartNumberWithColor(value, axis.numberFormat, chart.date1904, chart.indexedColors);
						return {
							value,
							label: formattedLabel.text,
							...formattedLabel.color ? { color: formattedLabel.color } : {},
							showLabel: !axis.deleted && labelPosition !== "none",
							x: horizontal ? position : lineCoordinate,
							y: horizontal ? lineCoordinate : position,
							labelX: horizontal ? position : lineCoordinate + (side === "l" ? -10 : 10),
							labelY: horizontal ? lineCoordinate + (side === "t" ? -10 : 22) : position + 5,
							labelAnchor: horizontal ? "middle" : side === "l" ? "end" : "start"
						};
					}),
					minorTicks: axis.minorTickMark && axis.minorTickMark !== "none" ? minorAxisTickValues(axisTickValues, minimum, upper, axis.minorUnit, axis.logBase).map((value) => {
						const position = tickPosition(value);
						return {
							x: horizontal ? position : lineCoordinate,
							y: horizontal ? lineCoordinate : position
						};
					}) : [],
					majorGridlines: !!axis.majorGridlines,
					...axis.majorTickMark ? { majorTickMark: axis.majorTickMark } : {},
					...axis.minorTickMark ? { minorTickMark: axis.minorTickMark } : {},
					...axis.deleted ? { deleted: true } : {}
				};
			});
			const xTickValues = axisMajorTickValues(chart.xAxis, xMinimum, xUpper);
			const xTickLabelPosition = chart.xAxis?.tickLabelPosition ?? "nextTo";
			const xTicks = xyChart ? xTickValues.map((value) => {
				const x = xPosition(value);
				const formattedLabel = formatChartNumberWithColor(value, chart.xAxis?.numberFormat, chart.date1904, chart.indexedColors);
				return {
					value,
					label: formattedLabel.text,
					...formattedLabel.color ? { color: formattedLabel.color } : {},
					showLabel: xTickLabelPosition !== "none" && !chart.xAxis?.deleted,
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
			if (!chart.valueAxis?.deleted) {
				if (radar) for (const tick of ticks) addAxisTick("value", chart.valueAxis?.majorTickMark, radarCenterX, tick.y, true, 1);
				else for (const tick of ticks) addAxisTick("value", chart.valueAxis?.majorTickMark, horizontalValueAxis ? tick.x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : tick.y, !horizontalValueAxis, valueTickInwardDirection);
				if (chart.valueAxis?.minorTickMark && chart.valueAxis.minorTickMark !== "none") for (const value of minorAxisTickValues(tickValues, minimum, upper, chart.valueAxis.minorUnit, chart.valueAxis.logBase)) {
					const x = radar ? radarCenterX : plot.left + plot.width * axisFraction(value);
					const y = radar ? radarCenterY - radarRadiusY * axisFraction(value) : valuePosition(value);
					if (radar) addAxisTick("value", chart.valueAxis.minorTickMark, radarCenterX, y, true, 1, void 0, true);
					else addAxisTick("value", chart.valueAxis.minorTickMark, horizontalValueAxis ? x : valueAxisLineCoordinate, horizontalValueAxis ? valueAxisLineCoordinate : y, !horizontalValueAxis, valueTickInwardDirection, void 0, true);
				}
			}
			if (xyChart && !chart.xAxis?.deleted) {
				const inwardDirection = xAxisSide === "t" ? 1 : -1;
				for (const tick of xTicks) addAxisTick("xValue", chart.xAxis?.majorTickMark, tick.x, tick.y, true, inwardDirection);
				if (chart.xAxis?.minorTickMark && chart.xAxis.minorTickMark !== "none") for (const value of minorAxisTickValues(xTickValues, xMinimum, xUpper, chart.xAxis.minorUnit, chart.xAxis.logBase)) addAxisTick("xValue", chart.xAxis.minorTickMark, xPosition(value), xyXAxisCrossingY, true, inwardDirection, void 0, true);
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
				else for (let index = 0; index < (dateAxisMajorTickValues?.length ?? categoryCount); index++) {
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
					const values = dateAxisMinorTickValues ? withoutMajorTickValues(dateAxisMinorTickValues, dateMinimum, dateMaximum, dateAxisMajorTickValues || []) : minorAxisTickValues(dateAxisMajorTickValues || dateValues || [], dateMinimum, dateMaximum);
					for (const value of values) {
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
				if (seriesType(series) !== "area") return;
				const axis = axisEntryForSeries(series)?.axis || chart.valueAxis;
				const { minimum, upper } = axisBoundsForSeries(series);
				const lowerValue = axis?.crossesAt ?? (axis?.crosses === "max" ? upper : axis?.crosses === "min" ? minimum : clamp(0, minimum, upper));
				let upperPoints = [];
				let lowerPoints = [];
				const flushArea = () => {
					if (upperPoints.length > 1) {
						const points = [...upperPoints, ...lowerPoints.reverse()];
						areas.push({
							d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z`,
							color: series.color,
							stroke: series.stroke ?? series.color,
							strokeWidth: chartStrokeWidth(series.strokeWidthEmu),
							...chartStrokeStyle(series, void 0, chartStrokeWidth(series.strokeWidthEmu) ?? 1)
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
					const stackedSegment = stackedSegmentFor(seriesIndex)?.[index];
					const centerX = categoryX(index);
					const centerY = categoryY(index);
					upperPoints.push(horizontalValueAxis ? {
						x: seriesValuePosition(series, stackedSegment?.end ?? value),
						y: centerY
					} : {
						x: centerX,
						y: seriesValuePosition(series, stackedSegment?.end ?? value)
					});
					lowerPoints.push(horizontalValueAxis ? {
						x: seriesValuePosition(series, stackedSegment?.start ?? lowerValue),
						y: centerY
					} : {
						x: centerX,
						y: seriesValuePosition(series, stackedSegment?.start ?? lowerValue)
					});
				}
				flushArea();
			});
			const formatValue = (value, series, index) => formatChartNumberWithColor(value, series.valueNumberFormatOverrides[index] ?? series.numberFormat ?? chart.numberFormat, chart.date1904, chart.indexedColors);
			const labelPosition = (series, index, fallback) => series.dataLabelPositionOverrides[index] ?? series.dataLabelPosition ?? chart.dataLabelPosition ?? fallback;
			const chartLegendKey = (series, index) => {
				const color = index === void 0 ? series.color : series.pointColors[index] || series.color;
				const type = seriesType(series);
				if (![
					"line",
					"scatter",
					"radar"
				].includes(type)) return { color };
				const symbol = index === void 0 ? series.markerSymbol : series.markerSymbolOverrides?.[index] ?? series.markerSymbol;
				const showMarker = symbol === void 0 ? type === "line" ? series.lineShowMarker !== false : type === "scatter" ? series.scatterShowMarker === true : series.radarShowMarker === true : symbol !== "none";
				const showLine = type === "scatter" ? series.scatterShowLine === true : type === "line" ? series.lineShowLine !== false : series.radarShowLine !== false;
				return {
					color,
					markerChart: true,
					...showMarker ? { symbol: symbol ?? "circle" } : {},
					...showLine ? { line: true } : {}
				};
			};
			const pushDataLabel = (label, x, y, anchor, series, index, leaderLine) => {
				const style = {
					...series.dataLabelStyle,
					...series.dataLabelStyleOverrides[index]
				};
				if (label.color) style.color = label.color;
				const lines = label.text.split(/\r?\n/);
				const showLegendKey = series.legendKeyOverrides[index] ?? series.showLegendKey;
				dataLabels.push({
					text: label.text,
					lines,
					x,
					y,
					anchor,
					...showLegendKey ? { legendKey: chartLegendKey(series, index) } : {},
					...leaderLine ? { leaderLine: {
						d: leaderLine,
						color: style.color || "#64748b"
					} } : {},
					...Object.keys(style).length ? { style } : {}
				});
			};
			const formatDataLabel = (series, index, value, category, total, bubbleSize) => {
				const showValue = series.valueLabelOverrides[index] ?? series.showValueLabels;
				const showLegendKey = series.legendKeyOverrides[index] ?? series.showLegendKey;
				const showCategory = series.categoryNameLabelOverrides[index] ?? series.showCategoryNameLabels;
				const showSeries = series.seriesNameLabelOverrides[index] ?? series.showSeriesNameLabels;
				const isPieLike = chart.type === "pie" || chart.type === "doughnut";
				const showPercent = isPieLike && (series.percentLabelOverrides[index] ?? series.showPercentLabels);
				const showBubbleSize = bubble && bubbleSize !== void 0 && (series.bubbleSizeLabelOverrides[index] ?? series.showBubbleSizeLabels);
				const parts = [
					showValue ? formatValue(value, series, index) : void 0,
					showCategory ? category : void 0,
					showSeries ? series.name : void 0,
					showPercent && total ? formatChartNumberWithColor(value / total, series.valueNumberFormatOverrides[index] ?? series.numberFormat ?? chart.numberFormat ?? "0.0%", chart.date1904, chart.indexedColors) : void 0,
					showBubbleSize ? formatValue(bubbleSize, series, index) : void 0
				].filter((part) => part !== void 0);
				if (!parts.length) return showLegendKey ? { text: "" } : void 0;
				const defaultSeparator = isPieLike && !showValue && showCategory && !showSeries && showPercent ? "\n" : ",";
				const color = parts.find((part) => typeof part !== "string" && Boolean(part.color));
				return {
					text: parts.map((part) => typeof part === "string" ? part : part.text).join(series.labelSeparatorOverrides[index] ?? series.labelSeparator ?? chart.labelSeparator ?? defaultSeparator),
					...color?.color ? { color: color.color } : {}
				};
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
					const volumeWidthFactor = 100 / (100 + (stock.volumeGapWidth ?? 150));
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
							color: volumeSeries.pointColors[index] || volumeSeries.color,
							stroke: volumeSeries.pointStrokes?.[index] ?? volumeSeries.stroke,
							strokeWidth: chartStrokeWidth(volumeSeries.pointStrokeWidthsEmu?.[index] ?? volumeSeries.strokeWidthEmu),
							...chartStrokeStyle(volumeSeries, index, chartStrokeWidth(volumeSeries.pointStrokeWidthsEmu?.[index] ?? volumeSeries.strokeWidthEmu) ?? 1)
						});
						const label = formatDataLabel(volumeSeries, index, value, chart.categories[index] || String(index + 1));
						if (label !== void 0) {
							const position = labelPosition(volumeSeries, index, "outEnd");
							let labelX = x;
							let labelY = value >= 0 ? top - 8 : bottom + 20;
							let anchor = "middle";
							if (position === "ctr") labelY = (zeroY + valueY) / 2 + 6;
							else if (position === "inBase") labelY = value >= 0 ? zeroY - 8 : zeroY + 18;
							else if (position === "inEnd") labelY = value >= 0 ? valueY + 18 : valueY - 4;
							else if (position === "b") labelY = bottom + 20;
							else if (position === "l") {
								labelX -= volumeWidth / 2 + 8;
								labelY = (zeroY + valueY) / 2 + 6;
								anchor = "end";
							} else if (position === "r") {
								labelX += volumeWidth / 2 + 8;
								labelY = (zeroY + valueY) / 2 + 6;
								anchor = "start";
							}
							pushDataLabel(label, labelX, labelY, anchor, volumeSeries, index);
						}
					}
				}
				for (let index = 0; index < chart.categories.length; index++) {
					const highValue = highSeries?.values[index];
					const lowValue = lowSeries?.values[index];
					const closeValue = closeSeries?.values[index];
					const openValue = openSeries?.values[index];
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
						const position = labelPosition(closeSeries, index, "r");
						pushDataLabel(label, position === "l" ? x - 10 : position === "r" ? x + 10 : x, position === "t" ? closeY - 12 : position === "b" ? closeY + 20 : closeY + 5, position === "l" ? "end" : position === "r" ? "start" : "middle", closeSeries, index);
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
					const categoryLabelY = categoryTickLabelPosition === "high" ? chart.valueAxis?.reverse ? plot.top + plot.height + 25 : plot.top - 10 : categoryTickLabelPosition === "low" ? chart.valueAxis?.reverse ? plot.top - 10 : plot.top + plot.height + 25 : categoryAxisSide === "t" ? categoryAxisCrossingCoordinate - 10 : categoryAxisCrossingCoordinate + 25;
					chart.categories.forEach((text, index) => {
						if (showCategoryAxisLabel(index)) categories.push({
							...formatCategoryAxisLabel(text),
							x: categoryX(index),
							y: categoryLabelY,
							anchor: "middle"
						});
					});
				}
				for (const index of [...volumeSeries && stock.volumeIndex !== void 0 ? [stock.volumeIndex] : [], ...priceIndices]) {
					const series = chart.series[index];
					if (series && !chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({
						text: series.name,
						x: 88 + legend.length * 220,
						y: 530,
						...chartLegendKey(series)
					});
				}
			} else if (chart.type === "area" && !combo) chart.series.forEach((series, seriesIndex) => {
				let upperPoints = [];
				let lowerPoints = [];
				const flushArea = () => {
					if (upperPoints.length > 1) {
						const points = [...upperPoints, ...lowerPoints.reverse()];
						areas.push({
							d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z`,
							color: series.color,
							stroke: series.stroke ?? series.color,
							strokeWidth: chartStrokeWidth(series.strokeWidthEmu),
							...chartStrokeStyle(series, void 0, chartStrokeWidth(series.strokeWidthEmu) ?? 1)
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
					const segment = stacking?.segments[seriesIndex]?.[index];
					const upperValue = stacking ? segment?.end ?? value : value;
					const lowerValue = stacking ? segment?.start ?? 0 : categoryAxisCrossingValue;
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
						const position = labelPosition(series, index, stacking ? "ctr" : "outEnd");
						const labelY = position === "ctr" ? (upperY + lowerY) / 2 + 5 : position === "inBase" ? lowerY - 8 : position === "inEnd" ? upperY + 18 : position === "b" ? Math.max(upperY, lowerY) + 20 : Math.min(upperY, lowerY) - 8;
						pushDataLabel(label, x, labelY, "middle", series, index);
					}
				}
				flushArea();
				if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({
					text: series.name,
					x: 88 + legend.length % 4 * 220,
					y: 530 + Math.floor(legend.length / 4) * 22,
					...chartLegendKey(series)
				});
			});
			else if (bubble) {
				const maximumSize = chart.series.flatMap((series) => (series.bubbleSizes || []).filter((size) => size !== null && size !== void 0 && size !== 0 && (size > 0 || chart.showNegativeBubbles === true))).reduce((maximum, size) => Math.max(maximum, Math.abs(size)), 0);
				const maximumRadius = Math.min(plot.width, plot.height) * .08 * (chart.bubbleScale ?? 100) / 100;
				chart.series.forEach((series, seriesIndex) => {
					const pointCount = Math.max(series.xValues?.length || 0, series.values.length, series.bubbleSizes?.length || 0);
					for (let index = 0; index < pointCount; index++) {
						const xValue = series.xValues?.[index];
						const yValue = series.values[index];
						const bubbleSize = series.bubbleSizes?.[index];
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
							stroke: series.pointStrokes?.[index] ?? series.stroke,
							strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
							...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
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
					if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({
						text: series.name,
						x: 88 + seriesIndex * 220,
						y: 530,
						...chartLegendKey(series)
					});
				});
			} else if (scatter) chart.series.forEach((series, seriesIndex) => {
				let path = "";
				let segment = [];
				const flushSegment = () => {
					path += smoothChartLinePath(segment, series.scatterSmooth === true);
					segment = [];
				};
				const pointCount = Math.max(series.xValues?.length || 0, series.values.length);
				for (let index = 0; index < pointCount; index++) {
					const xValue = series.xValues?.[index];
					const yValue = series.values[index];
					if (xValue == null || yValue == null) {
						flushSegment();
						continue;
					}
					const x = xPosition(xValue);
					const y = valuePosition(yValue);
					if (series.scatterShowLine) segment.push({
						x,
						y
					});
					const pointMarkerSymbol = series.markerSymbolOverrides?.[index] ?? series.markerSymbol;
					if (pointMarkerSymbol === void 0 ? series.scatterShowMarker : pointMarkerSymbol !== "none") markers.push({
						x,
						y,
						color: series.pointColors[index] || series.color,
						stroke: series.pointStrokes?.[index] ?? series.stroke,
						strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
						...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
						symbol: pointMarkerSymbol ?? "circle",
						radius: series.markerSizeOverrides?.[index] ?? series.markerSize
					});
					const label = formatDataLabel(series, index, yValue, String(xValue));
					if (label !== void 0) {
						const labelType = labelPosition(series, index, "r");
						const labelX = labelType === "l" ? x - 10 : labelType === "r" ? x + 10 : x;
						const labelY = labelType === "t" ? y - 12 : labelType === "b" ? y + 24 : y + 5;
						pushDataLabel(label, labelX, labelY, labelType === "l" ? "end" : labelType === "r" ? "start" : "middle", series, index);
					}
				}
				flushSegment();
				if (path) lines.push({
					d: path,
					color: series.stroke ?? series.color,
					width: chartStrokeWidth(series.strokeWidthEmu),
					...chartStrokeStyle(series, void 0, chartStrokeWidth(series.strokeWidthEmu) ?? 4)
				});
				if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({
					text: series.name,
					x: 88 + seriesIndex * 220,
					y: 530,
					...chartLegendKey(series)
				});
			});
			else if (chart.type === "pie" || chart.type === "doughnut") {
				const centerX = plot.left + plot.width / 2;
				const centerY = plot.top + plot.height / 2;
				const radius = Math.min(plot.width, plot.height) * .47;
				const doughnut = chart.type === "doughnut";
				const baseInnerRadius = doughnut ? radius * (chart.holeSize ?? 50) / 100 : 0;
				const ringSeries = doughnut ? chart.series : chart.series.slice(0, 1);
				const ringWidth = doughnut ? (radius - baseInnerRadius) / Math.max(ringSeries.length, 1) : radius;
				ringSeries.forEach((series, seriesIndex) => {
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
					let angle = -Math.PI / 2 + (chart.firstSliceAngle ?? 0) * Math.PI / 180;
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
						const strokeWidth = chartStrokeWidth(series.pointStrokeWidthsEmu?.[point.index] ?? series.strokeWidthEmu);
						slices.push({
							d,
							color: point.color,
							stroke: series.pointStrokes?.[point.index] ?? series.stroke ?? "white",
							strokeWidth: strokeWidth ?? 2,
							...chartStrokeStyle(series, point.index, strokeWidth ?? 2)
						});
						const label = formatDataLabel(series, point.index, point.value, point.text, total);
						if (label !== void 0) {
							const labelAngle = angle + sweep / 2;
							const position = labelPosition(series, point.index, "ctr");
							const hasLeaderLine = (series.leaderLineOverrides?.[point.index] ?? series.showLeaderLines ?? false) && position === "outEnd";
							const labelRadius = doughnut ? position === "outEnd" ? outerRadius * 1.12 : position === "inEnd" ? innerRadius + ringWidth * .82 : position === "inBase" ? innerRadius + ringWidth * .18 : (outerRadius + innerRadius) / 2 : radius * (position === "outEnd" ? 1.12 : position === "inEnd" ? .82 : position === "inBase" ? .24 : position === "bestFit" ? .62 : .68);
							const side = Math.cos(labelAngle) >= 0 ? 1 : -1;
							const leaderY = centerY + Math.sin(labelAngle) * outerRadius * 1.04;
							const x = hasLeaderLine ? centerX + side * outerRadius * 1.12 : position === "l" ? centerX - radius * .75 : position === "r" ? centerX + radius * .75 : centerX + Math.cos(labelAngle) * labelRadius;
							const y = hasLeaderLine ? leaderY + 6 : position === "t" ? centerY - radius * .78 : position === "b" ? centerY + radius * .78 : centerY + Math.sin(labelAngle) * labelRadius + 6;
							const anchor = hasLeaderLine ? side > 0 ? "start" : "end" : position === "l" ? "end" : position === "r" ? "start" : "middle";
							const leaderLine = hasLeaderLine ? `M ${centerX + Math.cos(labelAngle) * outerRadius} ${centerY + Math.sin(labelAngle) * outerRadius} L ${centerX + Math.cos(labelAngle) * outerRadius * 1.04} ${leaderY} L ${x - side * 6} ${leaderY}` : void 0;
							pushDataLabel(label, x, y, anchor, series, point.index, leaderLine);
						}
						angle = endAngle;
					}
					if (seriesIndex === 0) points.slice(0, 12).forEach((point) => {
						if (chart.legend?.hiddenEntries?.includes(point.index)) return;
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
				if (!chart.valueAxis?.deleted && chart.valueAxis?.majorGridlines) for (const fraction of [.../* @__PURE__ */ new Set([...radialValues, 1])].sort((a, b) => a - b)) {
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
						...formatCategoryAxisLabel(text),
						x: point.x,
						y: point.y + (Math.sin(point.angle) > .55 ? 16 : Math.sin(point.angle) < -0.55 ? -8 : 5),
						anchor: horizontal > .25 ? "start" : horizontal < -0.25 ? "end" : "middle"
					});
				});
				chart.series.forEach((series) => {
					const points = chart.categories.map((_category, index) => {
						const value = series.values[index];
						return value == null ? null : radarPoint(index, axisFraction(value));
					});
					const complete = points.length >= 3 && points.every((point) => point !== null);
					if (chart.radarStyle === "filled" && complete) areas.push({
						d: `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z`,
						color: series.color,
						stroke: series.stroke ?? series.color,
						strokeWidth: chartStrokeWidth(series.strokeWidthEmu),
						...chartStrokeStyle(series, void 0, chartStrokeWidth(series.strokeWidthEmu) ?? 1),
						opacity: .22
					});
					if (series.radarShowLine !== false) {
						const path = complete ? `${points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ")} Z` : points.flatMap((point, index) => {
							const next = points[(index + 1) % points.length];
							return point && next ? [`M ${point.x} ${point.y} L ${next.x} ${next.y}`] : [];
						}).join(" ");
						if (path) lines.push({
							d: path,
							color: series.stroke ?? series.color,
							width: chartStrokeWidth(series.strokeWidthEmu),
							...chartStrokeStyle(series, void 0, chartStrokeWidth(series.strokeWidthEmu) ?? 4)
						});
					}
					points.forEach((point, index) => {
						if (!point) return;
						const pointMarkerSymbol = series.markerSymbolOverrides?.[index] ?? series.markerSymbol;
						if (pointMarkerSymbol === void 0 ? series.radarShowMarker : pointMarkerSymbol !== "none") markers.push({
							x: point.x,
							y: point.y,
							color: series.pointColors[index] || series.color,
							stroke: series.pointStrokes?.[index] ?? series.stroke,
							strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
							...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
							symbol: pointMarkerSymbol ?? "circle",
							radius: series.markerSizeOverrides?.[index] ?? series.markerSize ?? series.radarMarkerSize ?? 5
						});
						const value = series.values[index];
						if (value == null) return;
						const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
						if (label === void 0) return;
						const position = labelPosition(series, index, "outEnd");
						const offset = position === "inBase" || position === "inEnd" || position === "ctr" ? -10 : 14;
						const x = position === "l" ? point.x - 12 : position === "r" ? point.x + 12 : point.x + Math.cos(point.angle) * offset;
						const y = position === "t" ? point.y - 14 : position === "b" ? point.y + 22 : point.y + Math.sin(point.angle) * offset + 6;
						const anchor = position === "l" ? "end" : position === "r" ? "start" : Math.cos(point.angle) > .25 ? "start" : Math.cos(point.angle) < -0.25 ? "end" : "middle";
						pushDataLabel(label, x, y, anchor, series, index);
					});
					if (!chart.legend?.hiddenEntries?.includes(series.legendIndex)) legend.push({
						text: series.name,
						x: 88 + legend.length * 220,
						y: 530,
						...chartLegendKey(series)
					});
				});
			} else {
				const barSeriesByGroup = /* @__PURE__ */ new Map();
				chart.series.forEach((series, seriesIndex) => {
					if (seriesType(series) !== "bar") return;
					const groupIndex = series.chartGroupIndex ?? 0;
					barSeriesByGroup.set(groupIndex, [...barSeriesByGroup.get(groupIndex) || [], seriesIndex]);
				});
				const linePointsBySeries = /* @__PURE__ */ new Map();
				chart.series.forEach((series, seriesIndex) => {
					if (seriesType(series) === "line" && series.lineShowLine !== false) linePointsBySeries.set(seriesIndex, []);
				});
				for (let index = 0; index < categoryCount; index++) {
					const centerX = categoryX(index);
					const centerY = categoryY(index);
					if (!chart.categoryAxisDeleted && categoryTickLabelPosition !== "none") {
						const categoryLabel = categoryTickLabelPosition === "high" ? horizontalValueAxis ? {
							x: chart.valueAxis?.reverse ? plot.left - 10 : plot.left + plot.width + 10,
							y: centerY + 5,
							anchor: chart.valueAxis?.reverse ? "end" : "start"
						} : {
							x: centerX,
							y: chart.valueAxis?.reverse ? plot.top + plot.height + 25 : plot.top - 10,
							anchor: "middle"
						} : categoryTickLabelPosition === "low" ? horizontalValueAxis ? {
							x: chart.valueAxis?.reverse ? plot.left + plot.width + 10 : plot.left - 10,
							y: centerY + 5,
							anchor: chart.valueAxis?.reverse ? "start" : "end"
						} : {
							x: centerX,
							y: chart.valueAxis?.reverse ? plot.top - 10 : plot.top + plot.height + 25,
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
							...formatCategoryAxisLabel(chart.categories[index] || ""),
							...categoryLabel
						});
					}
					chart.series.forEach((series, seriesIndex) => {
						const value = series.values[index];
						if (value === null || value === void 0) {
							linePointsBySeries.get(seriesIndex)?.push(null);
							return;
						}
						const color = series.color;
						const kind = seriesType(series);
						const stackedSegment = stackedSegmentFor(seriesIndex)?.[index];
						const isStacked = (series.grouping || chart.grouping || "clustered") !== "clustered";
						if (kind === "line") {
							const x = horizontalValueAxis ? seriesValuePosition(series, stackedSegment?.end ?? value) : centerX;
							const y = horizontalValueAxis ? centerY : seriesValuePosition(series, stackedSegment?.end ?? value);
							linePointsBySeries.get(seriesIndex)?.push({
								x,
								y
							});
							const pointMarkerSymbol = series.markerSymbolOverrides?.[index] ?? series.markerSymbol;
							if (pointMarkerSymbol === void 0 ? series.lineShowMarker !== false : pointMarkerSymbol !== "none") markers.push({
								x,
								y,
								color: series.pointColors[index] || color,
								stroke: series.pointStrokes?.[index] ?? series.stroke,
								strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
								...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1),
								symbol: pointMarkerSymbol ?? "circle",
								radius: series.markerSizeOverrides?.[index] ?? series.markerSize
							});
							const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
							if (label !== void 0) {
								const position = labelPosition(series, index, "t");
								const labelX = position === "l" ? x - 10 : position === "r" ? x + 10 : x;
								const labelY = position === "b" ? y + 24 : position === "ctr" || position === "inBase" || position === "inEnd" ? y + 7 : y - 12;
								pushDataLabel(label, labelX, labelY, position === "l" ? "end" : position === "r" ? "start" : "middle", series, index);
							}
						} else if (kind === "area") {
							const x = horizontalValueAxis ? seriesValuePosition(series, stackedSegment?.end ?? value) : centerX;
							const y = horizontalValueAxis ? centerY : seriesValuePosition(series, stackedSegment?.end ?? value);
							const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
							if (label !== void 0) {
								const position = labelPosition(series, index, "t");
								const labelX = position === "l" ? x - 10 : position === "r" ? x + 10 : x;
								const labelY = position === "b" ? y + 24 : position === "ctr" || position === "inBase" || position === "inEnd" ? y + 7 : y - 12;
								pushDataLabel(label, labelX, labelY, position === "l" ? "end" : position === "r" ? "start" : "middle", series, index);
							}
						} else if (chart.direction === "horizontal") {
							const groupIndexes = barSeriesByGroup.get(series.chartGroupIndex ?? 0) || [seriesIndex];
							const barOrdinal = groupIndexes.indexOf(seriesIndex);
							const seriesCount = Math.max(groupIndexes.length, 1);
							const { minimum: axisMinimum, upper: axisUpper } = axisBoundsForSeries(series);
							const groupHeight = categoryBandSize(index) * .72;
							const barHeight = groupHeight / (isStacked ? 1 : seriesCount);
							const barY = centerY - groupHeight / 2 + (isStacked ? 0 : barOrdinal * barHeight);
							const zeroX = seriesValuePosition(series, Math.max(axisMinimum, Math.min(axisUpper, 0)));
							const baseX = isStacked ? seriesValuePosition(series, stackedSegment?.start ?? 0) : zeroX;
							const valueX = seriesValuePosition(series, isStacked ? stackedSegment?.end ?? value : value);
							bars.push({
								x: Math.min(baseX, valueX),
								y: barY,
								width: Math.max(1, Math.abs(valueX - baseX)),
								height: barHeight * .86,
								color: series.pointColors[index] || color,
								stroke: series.pointStrokes?.[index] ?? series.stroke,
								strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
								...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1)
							});
							const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
							if (label !== void 0) {
								const position = labelPosition(series, index, isStacked ? "ctr" : "outEnd");
								const direction = value >= 0 ? 1 : -1;
								const left = Math.min(baseX, valueX);
								const right = Math.max(baseX, valueX);
								let labelX = valueX + direction * 8;
								let labelY = barY + barHeight * .68;
								let anchor = direction > 0 ? "start" : "end";
								if (position === "ctr") {
									labelX = (baseX + valueX) / 2;
									anchor = "middle";
								} else if (position === "inBase") {
									labelX = baseX + direction * 20;
									anchor = direction > 0 ? "start" : "end";
								} else if (position === "inEnd") {
									labelX = valueX - direction * 20;
									anchor = direction > 0 ? "end" : "start";
								} else if (position === "l") {
									labelX = left - 8;
									anchor = "end";
								} else if (position === "r") {
									labelX = right + 8;
									anchor = "start";
								} else if (position === "t") {
									labelX = (zeroX + valueX) / 2;
									labelY = barY - 8;
									anchor = "middle";
								} else if (position === "b") {
									labelX = (zeroX + valueX) / 2;
									labelY = barY + barHeight + 15;
									anchor = "middle";
								}
								pushDataLabel(label, labelX, labelY, anchor, series, index);
							}
						} else {
							const groupIndexes = barSeriesByGroup.get(series.chartGroupIndex ?? 0) || [seriesIndex];
							const barOrdinal = groupIndexes.indexOf(seriesIndex);
							const seriesCount = Math.max(groupIndexes.length, 1);
							const { minimum: axisMinimum, upper: axisUpper } = axisBoundsForSeries(series);
							const groupWidth = categoryBandSize(index) * .72;
							const barWidth = groupWidth / (isStacked ? 1 : seriesCount);
							const zeroY = seriesValuePosition(series, Math.max(axisMinimum, Math.min(axisUpper, 0)));
							const baseY = isStacked ? seriesValuePosition(series, stackedSegment?.start ?? 0) : zeroY;
							const valueY = seriesValuePosition(series, isStacked ? stackedSegment?.end ?? value : value);
							const barX = centerX - groupWidth / 2 + (isStacked ? 0 : barOrdinal * barWidth);
							bars.push({
								x: barX,
								y: Math.min(baseY, valueY),
								width: barWidth * .86,
								height: Math.max(1, Math.abs(valueY - baseY)),
								color: series.pointColors[index] || color,
								stroke: series.pointStrokes?.[index] ?? series.stroke,
								strokeWidth: chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu),
								...chartStrokeStyle(series, index, chartStrokeWidth(series.pointStrokeWidthsEmu?.[index] ?? series.strokeWidthEmu) ?? 1)
							});
							const label = formatDataLabel(series, index, value, chart.categories[index] || String(index + 1));
							if (label !== void 0) {
								const position = labelPosition(series, index, isStacked ? "ctr" : "outEnd");
								const top = Math.min(baseY, valueY);
								const bottom = Math.max(baseY, valueY);
								let labelX = barX + barWidth * .43;
								let labelY = value >= 0 ? top - 8 : bottom + 20;
								let anchor = "middle";
								if (position === "ctr") labelY = (baseY + valueY) / 2 + 6;
								else if (position === "inBase") labelY = value >= 0 ? baseY - 8 : baseY + 18;
								else if (position === "inEnd") labelY = value >= 0 ? valueY + 18 : valueY - 4;
								else if (position === "b") labelY = bottom + 20;
								else if (position === "l") {
									labelX = barX - 8;
									labelY = (baseY + valueY) / 2 + 6;
									anchor = "end";
								} else if (position === "r") {
									labelX = barX + barWidth * .86 + 8;
									labelY = (baseY + valueY) / 2 + 6;
									anchor = "start";
								}
								pushDataLabel(label, labelX, labelY, anchor, series, index);
							}
						}
					});
				}
				linePointsBySeries.forEach((points, seriesIndex) => {
					const series = chart.series[seriesIndex];
					if (!series) return;
					let path = "";
					let segment = [];
					const flushSegment = () => {
						path += smoothChartLinePath(segment, series.lineSmooth === true);
						segment = [];
					};
					points.forEach((point) => {
						if (point) segment.push(point);
						else flushSegment();
					});
					flushSegment();
					if (path) lines.push({
						d: path,
						color: series.stroke ?? series.color,
						width: chartStrokeWidth(series.strokeWidthEmu),
						...chartStrokeStyle(series, void 0, chartStrokeWidth(series.strokeWidthEmu) ?? 4)
					});
				});
				chart.series.slice(0, 12).forEach((series) => {
					if (chart.legend?.hiddenEntries?.includes(series.legendIndex)) return;
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
				const originX = manual?.x ? manual.x.mode === "edge" ? manual.x.value * 1e3 : defaultX + manual.x.value * 1e3 : defaultX;
				const originY = manual?.y ? manual.y.mode === "edge" ? manual.y.value * 600 : defaultY + manual.y.value * 600 : defaultY;
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
			const manualLayout = chart.legend?.manualLayout;
			let legendClip;
			if (manualLayout?.w || manualLayout?.h) {
				const x = manualLayout.x?.mode === "edge" ? manualLayout.x.value * 1e3 : defaultLegendOrigin(chart.legend?.position, "x") + (manualLayout.x?.value || 0) * 1e3;
				const y = manualLayout.y?.mode === "edge" ? manualLayout.y.value * 600 : defaultLegendOrigin(chart.legend?.position, "y") + (manualLayout.y?.value || 0) * 600;
				const width = manualLayout.w?.mode === "edge" ? manualLayout.w.value * 1e3 - x : (manualLayout.w?.value ?? 1) * 1e3;
				const height = manualLayout.h?.mode === "edge" ? manualLayout.h.value * 600 - y : (manualLayout.h?.value ?? 1) * 600;
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
			if (dateAxisMajorTickValues?.length && categories.length && hasDateAxisScale) {
				const firstLabel = categories[0];
				categories.splice(0, categories.length, ...dateAxisMajorTickValues.map((value) => {
					const fraction = dateAxisFraction(value);
					const formattedLabel = formatCategoryAxisLabel(String(value));
					return {
						...firstLabel,
						...formattedLabel,
						color: formattedLabel.color,
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
		watch(() => [
			props.slide.backgroundImage?.url,
			Boolean(props.slide.backgroundImage?.tile),
			props.slide.backgroundImage?.tile?.flip || "none",
			props.slide.backgroundImage?.crop.left,
			props.slide.backgroundImage?.crop.top,
			props.slide.backgroundImage?.crop.right,
			props.slide.backgroundImage?.crop.bottom
		], ([url, tiled, flip], _oldValue, onCleanup) => {
			backgroundImageSize.value = void 0;
			backgroundTileWarning.value = "";
		}, { immediate: true });
		const frameStyle = computed(() => {
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
			const naturalSize = backgroundImageSize.value?.sourceUrl === image.url ? backgroundImageSize.value : void 0;
			const tileWidth = naturalSize && tile ? naturalSize.width * tile.scaleX : 0;
			const tileHeight = naturalSize && tile ? naturalSize.height * tile.scaleY : 0;
			const tilePosition = (alignment, offset) => offset === 0 ? `${alignment}%` : `calc(${alignment}% ${offset < 0 ? "-" : "+"} ${Math.abs(offset)}px)`;
			const position = tile && tile.align === "tl" && tile.offsetX === 0 && tile.offsetY === 0 && (naturalSize?.repeatX ?? 1) === 1 && (naturalSize?.repeatY ?? 1) === 1 ? "0px 0px" : tile ? `${tilePosition(alignments[tile.align][0], tile.offsetX * scale.value + tileWidth * alignments[tile.align][0] / 100 * ((naturalSize?.repeatX ?? 1) - 1))} ${tilePosition(alignments[tile.align][1], tile.offsetY * scale.value + tileHeight * alignments[tile.align][1] / 100 * ((naturalSize?.repeatY ?? 1) - 1))}` : `${left + right ? left / (left + right) * 100 : 50}% ${top + bottom ? top / (top + bottom) * 100 : 50}%`;
			return {
				aspectRatio: `${props.width} / ${props.height}`,
				background: props.slide.background,
				backgroundImage: `url("${naturalSize?.url || image.url}")`,
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
		const textAnchorCenterOffsets = /* @__PURE__ */ new Map();
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
			const triggersByMedia = /* @__PURE__ */ new Map();
			props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
				sequence.conditions.forEach((condition, conditionIndex) => {
					if (condition.type !== "mediaBookmark") return;
					const triggers = triggersByMedia.get(condition.id) || [];
					triggers.push({
						sequenceIndex,
						conditionIndex,
						timeMs: condition.timeMs
					});
					triggersByMedia.set(condition.id, triggers);
				});
			});
			return triggersByMedia;
		});
		const executedMediaActions = /* @__PURE__ */ new Set();
		const scheduledMediaActions = /* @__PURE__ */ new Map();
		const startedMediaActions = /* @__PURE__ */ new Map();
		const completedMediaActions = reactive(/* @__PURE__ */ new Map());
		const endedMediaTargets = reactive(/* @__PURE__ */ new Map());
		const mediaActionEndTimers = /* @__PURE__ */ new Map();
		let automaticStartedAt;
		let stepStartedAt = /* @__PURE__ */ new Map();
		let mainStepEndedAt = /* @__PURE__ */ new Map();
		let triggerSteps = /* @__PURE__ */ new Map();
		let triggerStartedAt = /* @__PURE__ */ new Map();
		let triggerStepEndedAt = /* @__PURE__ */ new Map();
		let triggerPlaybackByKey = /* @__PURE__ */ new Map();
		const firedTimingNodeTriggers = /* @__PURE__ */ new Set();
		const triggerConditionDueAt = /* @__PURE__ */ new Map();
		const pendingTriggerConditions = /* @__PURE__ */ new Set();
		const triggerConditionKey = (sequenceIndex, conditionIndex) => `${sequenceIndex}:${conditionIndex}`;
		let autoFitFrame = 0;
		let animationFrame = 0;
		let animationTimer;
		const pixelRatio = ref(1);
		const scale = computed(() => stageWidth.value / Math.max(props.width, 1));
		const chartVisuals = computed(() => new Map(props.slide.elements.flatMap((element) => element.chart ? [[element.id, makeChartVisual(element.chart, element.id, element.width)]] : [])));
		const chartPatternDefinitions = computed(() => new Map(props.slide.elements.flatMap((element) => {
			if (!element.chart) return [];
			const fills = [...new Set(element.chart.series.flatMap((series) => [
				series.color,
				series.stroke,
				...Object.values(series.pointColors),
				...Object.values(series.pointStrokes || {})
			]).filter((fill) => fill !== void 0 && isPatternFillCss(fill)))];
			return fills.length ? [[element.id, fills.map((fill, index) => ({
				id: `${customGradientScope}-chart-pattern-${element.id.replace(/[^a-zA-Z0-9_-]/g, "_")}-${index}`,
				fill
			}))]] : [];
		})));
		const chartPaint = (elementId, fill) => {
			if (!fill || !isPatternFillCss(fill)) return fill;
			const pattern = chartPatternDefinitions.value.get(elementId)?.find((item) => item.fill === fill);
			return pattern ? `url(#${pattern.id})` : fill;
		};
		const chartDataLabelKeyPositions = ref(/* @__PURE__ */ new Map());
		const chartDataLabelKeyId = (elementId, index) => `${customGradientScope}:${elementId}:${index}`;
		const refreshChartDataLabelKeyPositions = async () => {
			await nextTick();
		};
		watch(chartVisuals, () => {
			refreshChartDataLabelKeyPositions();
		}, { flush: "post" });
		function supportsTextVerticalEllipsis(element) {
			return element.textVerticalOverflow === "ellipsis" && (element.textColumnCount ?? 1) === 1 && (!element.textOrientation || element.textOrientation === "horz");
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
			const visibleLines = bands.filter((line) => line.bottom <= contentBottom + .5).length;
			return bands.some((line) => line.bottom > contentBottom + .5) ? Math.max(1, visibleLines) : 0;
		}
		function scheduleShapeAutoFit() {
			if (!stage.value || autoFitFrame || scale.value <= 0) return;
			autoFitFrame = requestAnimationFrame(() => {
				autoFitFrame = 0;
				if (!stage.value) return;
				const ellipsisFrames = [...stage.value.querySelectorAll(".text-frame[data-text-overflow=\"ellipsis\"]")];
				let resetTextEllipsis = false;
				let layoutChanged = false;
				for (const frame of ellipsisFrames) {
					const id = frame.dataset.elementId;
					if (!id) continue;
					const deleted = textEllipsisLineCounts.delete(`${props.slide.id}:${id}`);
					layoutChanged ||= deleted;
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
				for (const frame of stage.value.querySelectorAll(".text-frame[data-autofit=\"shape\"]")) {
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
				for (const frame of stage.value.querySelectorAll(".text-frame[data-autofit=\"normal\"]")) {
					const id = frame.dataset.elementId;
					const element = id ? elements.get(id) : void 0;
					if (!id || !element || !frame.clientWidth || !frame.clientHeight) continue;
					const overflowWidth = frame.scrollWidth > frame.clientWidth + 1 ? frame.clientWidth / frame.scrollWidth : 1;
					const overflowHeight = frame.scrollHeight > frame.clientHeight + 1 ? frame.clientHeight / frame.scrollHeight : 1;
					if (overflowWidth === 1 && overflowHeight === 1) continue;
					const key = `${props.slide.id}:${id}`;
					const current = normalAutoFitScales.get(key) ?? element.textFontScale ?? 1;
					const fitted = Math.max(.05, current * Math.min(overflowWidth, overflowHeight) * .98);
					if (fitted < current - .001) {
						normalAutoFitScales.set(key, fitted);
						changed = true;
						layoutChanged = true;
					}
				}
				for (const cell of stage.value.querySelectorAll(".slide-table td[data-autofit=\"normal\"]")) {
					const id = cell.closest(".slide-element")?.dataset.elementId;
					const rowIndex = Number(cell.dataset.rowIndex);
					const cellIndex = Number(cell.dataset.cellIndex);
					const model = (id ? elements.get(id) : void 0)?.table?.rows[rowIndex]?.cells[cellIndex];
					const text = cell.querySelector(".table-cell-text");
					if (!id || !model || !text?.clientWidth || !text.clientHeight) continue;
					const overflowWidth = text.scrollWidth > text.clientWidth + 1 ? text.clientWidth / text.scrollWidth : 1;
					const overflowHeight = text.scrollHeight > text.clientHeight + 1 ? text.clientHeight / text.scrollHeight : 1;
					if (overflowWidth === 1 && overflowHeight === 1) continue;
					const key = tableCellAutoFitKey(id, rowIndex, cellIndex);
					const current = normalAutoFitScales.get(key) ?? model.textFontScale ?? 1;
					const fitted = Math.max(.05, current * Math.min(overflowWidth, overflowHeight) * .98);
					if (fitted < current - .001) {
						normalAutoFitScales.set(key, fitted);
						changed = true;
						layoutChanged = true;
					}
				}
				if (layoutChanged) elementLayoutVersion.value++;
				nextTick(measureTextAnchorCenterOffsets);
				if (changed) nextTick(scheduleShapeAutoFit);
			});
		}
		function measureTextAnchorCenterOffsets() {
			if (!stage.value) return;
			const elements = new Map(props.slide.elements.map((element) => [element.id, element]));
			let changed = false;
			for (const frame of stage.value.querySelectorAll(".text-frame")) {
				const id = frame.dataset.elementId;
				const element = id ? elements.get(id) : void 0;
				if (!element?.textAnchorCenter || (element.textColumnCount ?? 1) <= 1) continue;
				const bounds = frame.querySelector(".text-bounds");
				if (!bounds) continue;
				const frameRect = frame.getBoundingClientRect();
				const frameStyle = getComputedStyle(frame);
				const paddingLeft = Number.parseFloat(frameStyle.paddingLeft) || 0;
				const paddingRight = Number.parseFloat(frameStyle.paddingRight) || 0;
				const paddingTop = Number.parseFloat(frameStyle.paddingTop) || 0;
				const paddingBottom = Number.parseFloat(frameStyle.paddingBottom) || 0;
				const parent = frame.closest(".slide-element");
				const transform = parent ? getComputedStyle(parent).transform : "none";
				const matrix = transform === "none" ? new DOMMatrixReadOnly() : new DOMMatrixReadOnly(transform);
				const targetX = frameRect.left + frameRect.width / 2 + matrix.a * (paddingLeft - paddingRight) / 2 + matrix.c * (paddingTop - paddingBottom) / 2;
				const targetY = frameRect.top + frameRect.height / 2 + matrix.b * (paddingLeft - paddingRight) / 2 + matrix.d * (paddingTop - paddingBottom) / 2;
				const key = `${props.slide.id}:${id}`;
				const currentOffset = textAnchorCenterOffsets.get(key) ?? 0;
				const boundsRect = bounds.getBoundingClientRect();
				const offsetX = targetX - (boundsRect.left + boundsRect.width / 2 - matrix.a * currentOffset);
				const offsetY = targetY - (boundsRect.top + boundsRect.height / 2 - matrix.b * currentOffset);
				const axisScale = matrix.a * matrix.a + matrix.b * matrix.b;
				if (axisScale < 1e-6) continue;
				const nextOffset = (offsetX * matrix.a + offsetY * matrix.b) / axisScale;
				if (Math.abs(nextOffset) < .5) changed = textAnchorCenterOffsets.delete(key) || changed;
				else if (Math.abs(nextOffset - currentOffset) >= .5) {
					textAnchorCenterOffsets.set(key, nextOffset);
					changed = true;
				}
			}
			if (changed) elementLayoutVersion.value++;
		}
		const chartPlot = (id) => chartVisuals.value.get(id)?.plot ?? {
			left: 88,
			top: 75,
			width: 850,
			height: 390
		};
		const triggerKey = (sequence, step) => `${sequence}:${step}`;
		const timelineEnd = (action) => action.triggerSequence !== void 0 && action.triggerStep !== void 0 ? triggerStepEndedAt.get(triggerKey(action.triggerSequence, action.triggerStep)) : action.automatic ? void 0 : mainStepEndedAt.get(action.step);
		const timelineNow = (action) => Math.min(animationClock.value, timelineEnd(action) ?? Infinity);
		const timelineStart = (action) => {
			const startedAt = action.triggerSequence === void 0 || action.triggerStep === void 0 ? action.automatic ? automaticStartedAt : stepStartedAt.get(action.step) : triggerStartedAt.get(triggerKey(action.triggerSequence, action.triggerStep));
			if (startedAt === void 0 || startedAt === Number.NEGATIVE_INFINITY || !action.mediaWaitForEndKeys?.length) return startedAt;
			const endedAt = action.mediaWaitForEndKeys.map((key) => completedMediaActions.get(key));
			if (endedAt.some((time) => time === void 0)) return Infinity;
			return Math.max(startedAt, ...endedAt.map((time) => time + (action.mediaWaitDelayMs || 0) - action.delayMs));
		};
		const animationDuration = (action) => action.durationMs === 0 ? 0 : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (action.repeatCount ?? 1), action.repeatDurationMs ?? Infinity);
		const naturalAnimationDuration = (action) => action.effect === "media" ? action.mediaDurationMs ?? 0 : action.timingWarp ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs) : action.durationMs === 0 ? 0 : Math.min(action.durationMs * (action.autoReverse ? 2 : 1) * (Number.isFinite(action.repeatCount) || action.repeatDurationMs !== void 0 ? action.repeatCount ?? 1 : 1), action.repeatDurationMs ?? Infinity) + (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0);
		const animationScheduleDuration = (action) => action.effect === "media" ? 0 : action.timingWarp ? Math.max(0, action.timingWarp.groupEndDelayMs - action.delayMs) : animationDuration(action) + (action.iteration ? action.iteration.intervalMs * (action.iteration.ranges.length - 1) : 0);
		function iterationIndexFor(action, start, end) {
			const ranges = action.iteration?.ranges;
			if (!ranges) return void 0;
			let low = 0, high = ranges.length;
			while (low < high) {
				const middle = low + high >>> 1;
				if (ranges[middle].end <= start) low = middle + 1;
				else high = middle;
			}
			const range = ranges[low];
			return range && range.start <= start && range.end >= end ? action.iteration?.backwards ? ranges.length - low - 1 : low : -1;
		}
		const animationsByTarget = computed(() => {
			const targets = /* @__PURE__ */ new Map();
			if (!props.showAnimations) return targets;
			for (const animation of props.slide.automaticAnimations || []) {
				const actions = targets.get(animation.targetId) || [];
				actions.push({
					...animation,
					step: -1,
					automatic: true
				});
				targets.set(animation.targetId, actions);
			}
			props.slide.animationSteps?.forEach((step, index) => {
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
			props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
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
		const playbackActionEnabled = (action) => (props.showAnimations || action.effect === "media") && (props.showNarration || !action.mediaIsNarration);
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
		].filter(playbackActionEnabled));
		const triggerSequenceHasEnabledAction = (sequence) => sequence.steps.some((step) => step.some(playbackActionEnabled));
		const triggerTargetIds = computed(() => {
			const targetIds = /* @__PURE__ */ new Set();
			if (props.thumbnail) return targetIds;
			for (const sequence of props.slide.triggeredAnimations || []) {
				if (!triggerSequenceHasEnabledAction(sequence)) continue;
				for (const condition of sequence.conditions) if (condition.type === "shape") targetIds.add(condition.id);
			}
			return targetIds;
		});
		function endTimelineStep(triggerSequence, step, endedAt) {
			if (triggerSequence === void 0) mainStepEndedAt.set(step, endedAt);
			else triggerStepEndedAt.set(triggerKey(triggerSequence, step), endedAt);
			timelineActions.value.forEach((action, index) => {
				if (!(triggerSequence === void 0 ? action.triggerSequence === void 0 && !action.automatic && action.step === step : action.triggerSequence === triggerSequence && action.triggerStep === step) || action.effect !== "media" || !executedMediaActions.has(index)) return;
				const media = Array.from(stage.value?.querySelectorAll("[data-media-id]") || []).find((item) => item.dataset.mediaId === action.targetId);
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
			runTimedMediaActions();
			const nextTimingNodeTriggerAt = runTimingNodeTriggeredAnimations(animationClock.value);
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
				const end = Math.min(start + animationScheduleDuration(action), stoppedAt ?? Infinity);
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
			scheduledMediaActions.clear();
			startedMediaActions.clear();
			completedMediaActions.clear();
			endedMediaTargets.clear();
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
			triggerConditionDueAt.clear();
			pendingTriggerConditions.clear();
			animationClock.value = now;
			scheduleTimelineUpdate();
		}
		function syncTriggerPlayback() {
			if (props.triggerPlayback === void 0) return;
			triggerSteps.clear();
			triggerStartedAt.clear();
			triggerStepEndedAt.clear();
			triggerPlaybackByKey.clear();
			for (const playback of [...props.triggerPlayback].sort((a, b) => a.startedAt - b.startedAt)) {
				const sequence = props.slide.triggeredAnimations?.[playback.sequence];
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
			if (!stage.value) return;
			if (slideId !== oldSlideId) {
				autoFitDimensions.clear();
				normalAutoFitScales.clear();
				textEllipsisLineCounts.clear();
				textAnchorCenterOffsets.clear();
				beginSlideTimeline(step);
				nextTick(scheduleShapeAutoFit);
				return;
			}
			if (step > oldStep) {
				for (let previous = oldStep + 1; previous < step; previous++) stepStartedAt.set(previous, Number.NEGATIVE_INFINITY);
				if (step === oldStep + 1) runSlideTriggeredAnimations("onNext");
				const now = performance.now();
				const previousStep = oldStep - 1;
				if (previousStep >= 0 && !props.slide.animationSequence?.concurrent && props.slide.animationSequence?.nextAction !== "seek") endTimelineStep(void 0, previousStep, now);
				if (props.slide.animationSequence?.nextAction === "seek" && oldStep > 0) {
					const startedAt = stepStartedAt.get(previousStep);
					const naturalEnd = Math.max(0, ...(props.slide.animationSteps?.[previousStep] || []).map((action) => action.delayMs + naturalAnimationDuration(action)));
					if (startedAt !== void 0 && startedAt !== Number.NEGATIVE_INFINITY) stepStartedAt.set(previousStep, Math.min(startedAt, now - naturalEnd));
				}
				stepStartedAt.set(oldStep, now);
			} else if (step < oldStep) {
				if (step === oldStep - 1) runSlideTriggeredAnimations("onPrev");
				props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
					sequence.conditions.forEach((condition, conditionIndex) => {
						if (condition.type !== "timeNode" || condition.sourceStep === void 0 || condition.sourceStep < step) return;
						const key = triggerConditionKey(sequenceIndex, conditionIndex);
						firedTimingNodeTriggers.delete(key);
						triggerConditionDueAt.delete(key);
						pendingTriggerConditions.delete(key);
					});
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
			const fittedSize = autoFitDimensions.get(`${props.slide.id}:${element.id}`);
			const width = fittedSize?.width || element.width;
			const height = fittedSize?.height || element.height;
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
			const shadows = element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : [];
			const fillOn = animationOnOff(element, "fill.on");
			const strokeOn = animationOnOff(element, "stroke.on");
			const shadowOn = animationOnOff(element, "shadow.on");
			const shadowBlurScale = element.kind === "shape" && element.geometry === "ellipse" ? .5 : 1;
			const shadowColors = shadows.map((shadow) => shadowCss(shadow, applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity")), shadow.inset ? 1 : shadowBlurScale));
			const animatedFill = animationColor(element, "fillcolor");
			const fill = animatedFill || element.fill;
			const fillOpacity = animationOpacity(element, "fill.opacity");
			const stroke = applyColorOpacity(animationColor(element, "stroke.color") || element.stroke, animationOpacity(element, "stroke.opacity"));
			const strokeWidth = animationStrokeWidth(element);
			const strokeDash = animationStrokeDash(element);
			const glows = element.glows?.length ? element.glows : element.glow ? [element.glow] : [];
			const customGeometry = Boolean(element.customPaths?.length);
			const radialFill = fillOn && !animatedFill && !customGeometry && !usesSvgPolygon(element) ? customSvgGradient(element, "fill") : void 0;
			const radialBackground = radialFill?.kind === "radial" && radialFill.fx !== void 0 && radialFill.fy !== void 0 && radialFill.fr !== void 0 ? radialGradientBackground(radialFill) : void 0;
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
			return element.kind === "shape" && isSvgPolygonGeometry(element.geometry);
		}
		function presetBracketBracePaths(element) {
			const rectangle = "M0 0 H1000 V1000 H0 Z";
			if (!(element.width > 0) || !(element.height > 0)) return [rectangle, rectangle];
			const shortSide = Math.min(element.width, element.height);
			if (element.geometry === "leftBrace" || element.geometry === "rightBrace") {
				const adjustment2 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj2 ?? 5e4));
				const maximumAdjustment1 = Math.min(adjustment2, 1e5 - adjustment2) / 2 * element.height / shortSide;
				const radiusY = shortSide * Math.max(0, Math.min(maximumAdjustment1, element.presetAdjustments?.adj1 ?? 8333)) / 1e5;
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
				const adjustment = Math.max(0, Math.min(25e3, element.presetAdjustments?.adj ?? 8333));
				const radius = shortSide * adjustment / 1e5;
				const x2 = shortSide * adjustment / 5e4;
				const x3 = element.width - x2, x4 = element.width - radius;
				const y2 = element.height / 2 - radius, y3 = element.height / 2 + radius, y4 = element.height - radius;
				const point = ([x, y]) => `${Number((x / element.width * 1e3).toFixed(6))} ${Number((y / element.height * 1e3).toFixed(6))}`;
				const buildPath = (outline) => {
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
					return (outline ? [...leftBrace, ...rightBrace] : [
						...leftBrace,
						joinAtTop,
						...rightBrace.slice(1),
						"Z"
					]).filter(Boolean).join(" ");
				};
				return [buildPath(false), buildPath(true)];
			}
			const maximumAdjustment = isPair ? 5e4 : 5e4 * element.height / shortSide;
			const adjustment = Math.max(0, Math.min(maximumAdjustment, element.presetAdjustments?.adj ?? (isPair ? 16667 : 8333)));
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
			if (element.geometry === "leftBracket" || element.geometry === "rightBracket" || element.geometry === "bracketPair" || element.geometry === "leftBrace" || element.geometry === "rightBrace" || element.geometry === "bracePair") return [presetBracketBracePaths(element)[1]];
			if (element.geometry === "arc") {
				const arc = presetAngleArcPaths(element)[1];
				return arc ? [arc] : [];
			}
			if (element.geometry === "frame") return [presetOutlinePath(element)];
			if (element.geometry === "cube") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return [presetOutlinePath(element)];
				const inset = Math.min(width, height) * (element.presetAdjustment ?? .25);
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
				const scale = 1e3 / 21600;
				const point = (x, y) => `${x * scale} ${y * scale}`;
				return [
					`M${point(0, 3675)} L${point(18595, 3675)} L${point(18595, 18022)} C${point(9298, 18022)} ${point(9298, 23542)} ${point(0, 20782)} Z`,
					`M${point(1532, 3675)} L${point(1532, 1815)} L${point(2e4, 1815)} L${point(2e4, 16252)} C${point(19298, 16252)} ${point(18595, 16352)} ${point(18595, 16352)}`,
					`M${point(2972, 1815)} L${point(2972, 0)} L${point(21600, 0)} L${point(21600, 14392)} C${point(20800, 14392)} ${point(2e4, 14467)} ${point(2e4, 14467)}`
				];
			}
			return [];
		}
		function roundRectRadii(element) {
			const fitted = autoFitDimensions.get(`${props.slide.id}:${element.id}`);
			const width = fitted?.width || element.width, height = fitted?.height || element.height;
			const radius = Math.min(width, height) * (element.presetAdjustment ?? .16667);
			return [width > 0 ? radius / width * 1e3 : 0, height > 0 ? radius / height * 1e3 : 0];
		}
		function presetAngleArcPaths(element) {
			const width = element.width, height = element.height;
			if (width <= 0 || height <= 0) return ["M0 0 H1000 V1000 H0 Z", ""];
			const fullTurn = 216e5;
			const startAngle = element.presetAdjustments?.adj1 ?? (element.geometry === "pie" ? 0 : element.geometry === "chord" ? 27e5 : 162e5);
			const endAngle = element.presetAdjustments?.adj2 ?? (element.geometry === "arc" ? 0 : 162e5);
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
			const width = element.width, height = element.height;
			if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
			const fullTurn = 216e5;
			const startAngle = element.presetAdjustments?.adj1 ?? 108e5;
			const endAngle = element.presetAdjustments?.adj2 ?? 0;
			let sweepAngle = endAngle - startAngle;
			if (sweepAngle <= 0) sweepAngle += fullTurn;
			const inset = Math.min(width, height) * (element.presetAdjustments?.adj3 ?? 25e3) / 1e5;
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
			if (element.geometry === "leftBracket" || element.geometry === "rightBracket" || element.geometry === "bracketPair" || element.geometry === "leftBrace" || element.geometry === "rightBrace" || element.geometry === "bracePair") return presetBracketBracePaths(element)[0];
			if (element.geometry === "quadArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj2 = Math.max(0, Math.min(5e4, element.presetAdjustments?.adj2 ?? 22500));
				const adj1 = Math.max(0, Math.min(2 * adj2, element.presetAdjustments?.adj1 ?? 22500));
				const x1 = shortSide * Math.max(0, Math.min((1e5 - 2 * adj2) / 2, element.presetAdjustments?.adj3 ?? 22500)) / 1e5;
				const dx2 = shortSide * adj2 / 1e5;
				const dx3 = shortSide * adj1 / 2e5;
				const x2 = width / 2 - dx2, x3 = width / 2 - dx3, x4 = width / 2 + dx3, x5 = width / 2 + dx2, x6 = width - x1;
				const y2 = height / 2 - dx2, y3 = height / 2 - dx3, y4 = height / 2 + dx3, y5 = height / 2 + dx2, y6 = height - x1;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 500 L${x(x1)} ${y(y2)} L${x(x1)} ${y(y3)} L${x(x3)} ${y(y3)} L${x(x3)} ${y(x1)} L${x(x2)} ${y(x1)} L500 0 L${x(x5)} ${y(x1)} L${x(x4)} ${y(x1)} L${x(x4)} ${y(y3)} L${x(x6)} ${y(y3)} L${x(x6)} ${y(y2)} L1000 500 L${x(x6)} ${y(y5)} L${x(x6)} ${y(y4)} L${x(x4)} ${y(y4)} L${x(x4)} ${y(y6)} L${x(x5)} ${y(y6)} L500 1000 L${x(x2)} ${y(y6)} L${x(x3)} ${y(y6)} L${x(x3)} ${y(y4)} L${x(x1)} ${y(y4)} L${x(x1)} ${y(y5)} Z`;
			}
			if (element.geometry === "leftRightUpArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj2 = Math.max(0, Math.min(5e4, element.presetAdjustments?.adj2 ?? 25e3));
				const adj1 = Math.max(0, Math.min(2 * adj2, element.presetAdjustments?.adj1 ?? 25e3));
				const x1 = shortSide * Math.max(0, Math.min((1e5 - 2 * adj2) / 2, element.presetAdjustments?.adj3 ?? 25e3)) / 1e5;
				const dx2 = shortSide * adj2 / 1e5;
				const dx3 = shortSide * adj1 / 2e5;
				const x2 = width / 2 - dx2, x3 = width / 2 - dx3, x4 = width / 2 + dx3, x5 = width / 2 + dx2, x6 = width - x1;
				const y2 = height - shortSide * adj2 / 5e4, y4 = height - dx2, y3 = y4 - dx3, y5 = y4 + dx3;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 ${y(y4)} L${x(x1)} ${y(y2)} L${x(x1)} ${y(y3)} L${x(x3)} ${y(y3)} L${x(x3)} ${y(x1)} L${x(x2)} ${y(x1)} L500 0 L${x(x5)} ${y(x1)} L${x(x4)} ${y(x1)} L${x(x4)} ${y(y3)} L${x(x6)} ${y(y3)} L${x(x6)} ${y(y2)} L1000 ${y(y4)} L${x(x6)} 1000 L${x(x6)} ${y(y5)} L${x(x1)} ${y(y5)} L${x(x1)} 1000 Z`;
			}
			if (element.geometry === "bentUpArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(5e4, element.presetAdjustments?.adj1 ?? 25e3));
				const adj2 = Math.max(0, Math.min(5e4, element.presetAdjustments?.adj2 ?? 25e3));
				const adj3 = Math.max(0, Math.min(5e4, element.presetAdjustments?.adj3 ?? 25e3));
				const dx2 = shortSide * adj1 / 2e5;
				const x3 = width - shortSide * adj1 / 1e5;
				const x2 = x3 - dx2, x4 = x3 + dx2;
				const y1 = shortSide * adj3 / 1e5;
				const y2 = height - shortSide * adj1 / 1e5;
				const x1 = width - shortSide * adj2 / 5e4;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 ${y(y2)} L${x(x2)} ${y(y2)} L${x(x2)} ${y(y1)} L${x(x1)} ${y(y1)} L${x(x3)} 0 L1000 ${y(y1)} L${x(x4)} ${y(y1)} V1000 H0 Z`;
			}
			if (element.geometry === "upArrow" || element.geometry === "downArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 5e4));
				const adj2 = Math.max(0, Math.min(1e5 * height / shortSide, element.presetAdjustments?.adj2 ?? 5e4));
				const halfHeadWidth = width * adj1 / 2e5;
				const x1 = width / 2 - halfHeadWidth, x2 = width / 2 + halfHeadWidth;
				const arrowDepth = shortSide * adj2 / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				if (element.geometry === "upArrow") {
					const y2 = arrowDepth;
					return `M0 ${y(y2)} L500 0 L1000 ${y(y2)} L${x(x2)} ${y(y2)} V1000 H${x(x1)} V${y(y2)} Z`;
				}
				const y1 = height - arrowDepth;
				return `M0 ${y(y1)} L${x(x1)} ${y(y1)} V0 H${x(x2)} V${y(y1)} H1000 L500 1000 Z`;
			}
			if (element.geometry === "leftArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 5e4));
				const x2 = shortSide * Math.max(0, Math.min(1e5 * width / shortSide, element.presetAdjustments?.adj2 ?? 5e4)) / 1e5;
				const arrowHalfHeight = height * adj1 / 2e5;
				const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 500 L${x(x2)} 0 L${x(x2)} ${y(y1)} H1000 V${y(y2)} H${x(x2)} L${x(x2)} 1000 Z`;
			}
			if (element.geometry === "notchedRightArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 5e4));
				const arrowLength = shortSide * Math.max(0, Math.min(1e5 * width / shortSide, element.presetAdjustments?.adj2 ?? 5e4)) / 1e5;
				const x2 = width - arrowLength;
				const arrowHalfHeight = height * adj1 / 2e5;
				const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
				const x1 = arrowHalfHeight * arrowLength / (height / 2);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 ${y(y1)} L${x(x2)} ${y(y1)} L${x(x2)} 0 L1000 500 L${x(x2)} 1000 L${x(x2)} ${y(y2)} H0 L${x(x1)} 500 Z`;
			}
			if (element.geometry === "rightArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 5e4));
				const x1 = width - shortSide * Math.max(0, Math.min(1e5 * width / shortSide, element.presetAdjustments?.adj2 ?? 5e4)) / 1e5;
				const arrowHalfHeight = height * adj1 / 2e5;
				const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 ${y(y1)} L${x(x1)} ${y(y1)} L${x(x1)} 0 L1000 500 L${x(x1)} 1000 L${x(x1)} ${y(y2)} H0 Z`;
			}
			if (element.geometry === "upDownArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 5e4));
				const adj2 = Math.max(0, Math.min(5e4 * height / shortSide, element.presetAdjustments?.adj2 ?? 5e4));
				const arrowHalfWidth = width * adj1 / 2e5;
				const x1 = width / 2 - arrowHalfWidth, x2 = width / 2 + arrowHalfWidth;
				const arrowHeight = shortSide * adj2 / 1e5;
				const y2 = arrowHeight, y3 = height - arrowHeight;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M500 0 L1000 ${y(y2)} L${x(x2)} ${y(y2)} L${x(x2)} ${y(y3)} L1000 ${y(y3)} L500 1000 L0 ${y(y3)} L${x(x1)} ${y(y3)} L${x(x1)} ${y(y2)} L0 ${y(y2)} Z`;
			}
			if (element.geometry === "leftRightArrow") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 5e4));
				const arrowWidth = shortSide * Math.max(0, Math.min(5e4 * width / shortSide, element.presetAdjustments?.adj2 ?? 5e4)) / 1e5;
				const x2 = arrowWidth, x3 = width - arrowWidth;
				const arrowHalfHeight = height * adj1 / 2e5;
				const y1 = height / 2 - arrowHalfHeight, y2 = height / 2 + arrowHalfHeight;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 500 L${x(x2)} 0 L${x(x2)} ${y(y1)} L${x(x3)} ${y(y1)} L${x(x3)} 0 L1000 500 L${x(x3)} 1000 L${x(x3)} ${y(y2)} L${x(x2)} ${y(y2)} L${x(x2)} 1000 Z`;
			}
			if (element.geometry === "plus") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const inset = Math.min(width, height) * Math.max(0, Math.min(.5, element.presetAdjustment ?? .25));
				const x1 = inset / width * 1e3, x2 = 1e3 - x1;
				const y1 = inset / height * 1e3, y2 = 1e3 - y1;
				return `M0 ${y1} L${x1} ${y1} L${x1} 0 L${x2} 0 L${x2} ${y1} H1000 V${y2} H${x2} V1000 H${x1} V${y2} H0 Z`;
			}
			if (element.geometry === "mathMinus") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const adj1 = Math.max(0, Math.min(1e5, element.presetAdjustments?.adj1 ?? 23520));
				const halfBarWidth = width * 73490 / 2e5;
				const halfBarHeight = height * adj1 / 2e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(width / 2 - halfBarWidth)} ${y(height / 2 - halfBarHeight)} H${x(width / 2 + halfBarWidth)} V${y(height / 2 + halfBarHeight)} H${x(width / 2 - halfBarWidth)} Z`;
			}
			if (element.geometry === "mathEqual") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const adj1 = Math.max(0, Math.min(36745, element.presetAdjustments?.adj1 ?? 23520));
				const adj2 = Math.max(0, Math.min(1e5 - 2 * adj1, element.presetAdjustments?.adj2 ?? 11760));
				const halfBarWidth = width * 73490 / 2e5;
				const halfGap = height * adj2 / 2e5;
				const barHeight = height * adj1 / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				const left = x(width / 2 - halfBarWidth), right = x(width / 2 + halfBarWidth);
				const upperInner = y(height / 2 - halfGap);
				return `M${left} ${y(height / 2 - halfGap - barHeight)} H${right} V${upperInner} H${left} Z M${left} ${y(height / 2 + halfGap)} H${right} V${y(height / 2 + halfGap + barHeight)} H${left} Z`;
			}
			if (element.geometry === "mathPlus") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const adj1 = Math.max(0, Math.min(73490, element.presetAdjustments?.adj1 ?? 23520));
				const longHalfArm = width * 73490 / 2e5;
				const crossHalfArm = Math.min(width, height) * adj1 / 2e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				const x1 = x(width / 2 - longHalfArm), x2 = x(width / 2 - crossHalfArm), x3 = x(width / 2 + crossHalfArm), x4 = x(width / 2 + longHalfArm);
				const y1 = y(height / 2 - height * 73490 / 2e5), y2 = y(height / 2 - crossHalfArm), y3 = y(height / 2 + crossHalfArm), y4 = y(height / 2 + height * 73490 / 2e5);
				return `M${x1} ${y2} L${x2} ${y2} L${x2} ${y1} L${x3} ${y1} L${x3} ${y2} L${x4} ${y2} L${x4} ${y3} L${x3} ${y3} L${x3} ${y4} L${x2} ${y4} L${x2} ${y3} L${x1} ${y3} Z`;
			}
			if (element.geometry === "mathDivide") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const adj1 = Math.max(1e3, Math.min(36745, element.presetAdjustments?.adj1 ?? 23520));
				const maximumAdj3 = Math.min((73490 + adj1) / 4, 36745 * width / height);
				const adj3 = Math.max(1e3, Math.min(maximumAdj3, element.presetAdjustments?.adj3 ?? 11760));
				const maximumAdj2 = 73490 - 4 * adj3 - adj1;
				const adj2 = Math.max(0, Math.min(maximumAdj2, element.presetAdjustments?.adj2 ?? 5880));
				const dy1 = height * adj1 / 2e5, gap = height * adj2 / 1e5, radius = height * adj3 / 1e5;
				const y3 = height / 2 - dy1, y4 = height / 2 + dy1;
				const y1 = y3 - gap, y2 = y1 + radius, y5 = height - y1;
				const halfBarWidth = width * 73490 / 2e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				const circle = (centerY) => `M${x(width / 2)} ${y(centerY - radius)} A${x(radius)} ${y(radius)} 0 1 1 ${x(width / 2)} ${y(centerY + radius)} A${x(radius)} ${y(radius)} 0 1 1 ${x(width / 2)} ${y(centerY - radius)} Z`;
				return `${circle(y2)} ${circle(y5 - radius)} M${x(width / 2 - halfBarWidth)} ${y(y3)} H${x(width / 2 + halfBarWidth)} V${y(y4)} H${x(width / 2 - halfBarWidth)} Z`;
			}
			if (element.geometry === "mathMultiply") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const adj1 = Math.max(0, Math.min(51965, element.presetAdjustments?.adj1 ?? 23520));
				const angle = Math.atan2(height, width), tangent = Math.tan(angle);
				const widthAlongDiagonal = Math.hypot(width, height) * (1 - .51965);
				const thickness = Math.min(width, height) * adj1 / 1e5;
				const xM = Math.cos(angle) * widthAlongDiagonal / 2, yM = Math.sin(angle) * widthAlongDiagonal / 2;
				const dx = Math.sin(angle) * thickness / 2, dy = Math.cos(angle) * thickness / 2;
				const xA = xM - dx, yA = yM + dy, xB = xM + dx, yB = yM - dy;
				const yC = (width / 2 - xB) * tangent + yB;
				const xD = width - xB, xE = width - xA;
				const xFE = tangent ? (height / 2 - yA) / tangent : 0;
				const xF = xE - xFE, xL = xA + xFE;
				const yG = height - yA, yH = height - yB, yI = height - yC;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(xA)} ${y(yA)} L${x(xB)} ${y(yB)} L500 ${y(yC)} L${x(xD)} ${y(yB)} L${x(xE)} ${y(yA)} L${x(xF)} 500 L${x(xE)} ${y(yG)} L${x(xD)} ${y(yH)} L500 ${y(yI)} L${x(xB)} ${y(yH)} L${x(xA)} ${y(yG)} L${x(xL)} 500 Z`;
			}
			if (element.geometry === "mathNotEqual") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const adj1 = Math.max(0, Math.min(5e4, element.presetAdjustments?.adj1 ?? 23520));
				const angle = Math.max(42e5, Math.min(66e5, element.presetAdjustments?.adj2 ?? 66e5)) - 54e5;
				const adj3 = Math.max(0, Math.min(1e5 - 2 * adj1, element.presetAdjustments?.adj3 ?? 11760));
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
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(x1)} ${y(y1)} L${x(x6)} ${y(y1)} L${x(lx)} ${y(ly)} L${x(rx)} ${y(ry)} L${x(rx6)} ${y(y1)} L${x(x8)} ${y(y1)} L${x(x8)} ${y(y2)} L${x(rx5)} ${y(y2)} L${x(rx4)} ${y(y3)} L${x(x8)} ${y(y3)} L${x(x8)} ${y(y4)} L${x(rx3)} ${y(y4)} L${x(drx)} ${y(dry)} L${x(dlx)} ${y(dly)} L${x(x3)} ${y(y4)} L${x(x1)} ${y(y4)} L${x(x1)} ${y(y3)} L${x(x4)} ${y(y3)} L${x(x5)} ${y(y2)} L${x(x1)} ${y(y2)} Z`;
			}
			if (element.geometry === "snip1Rect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const inset = Math.min(width, height) * (element.presetAdjustment ?? .16667);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 0 H${x(width - inset)} L1000 ${y(inset)} V1000 H0 Z`;
			}
			if (element.geometry === "snip2DiagRect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const leftCut = shortSide * Math.max(0, Math.min(5e4, element.presetAdjustments?.adj1 ?? 0)) / 1e5;
				const rightCut = shortSide * Math.max(0, Math.min(5e4, element.presetAdjustments?.adj2 ?? 16667)) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(leftCut)} 0 L${x(width - rightCut)} 0 L1000 ${y(rightCut)} L1000 ${y(height - leftCut)} L${x(width - leftCut)} 1000 L${x(rightCut)} 1000 L0 ${y(height - rightCut)} L0 ${y(leftCut)} Z`;
			}
			if (element.geometry === "snip2SameRect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const topCut = shortSide * Math.max(0, Math.min(5e4, element.presetAdjustments?.adj1 ?? 16667)) / 1e5;
				const bottomCut = shortSide * Math.max(0, Math.min(5e4, element.presetAdjustments?.adj2 ?? 0)) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(topCut)} 0 L${x(width - topCut)} 0 L1000 ${y(topCut)} L1000 ${y(height - bottomCut)} L${x(width - bottomCut)} 1000 L${x(bottomCut)} 1000 L0 ${y(height - bottomCut)} L0 ${y(topCut)} Z`;
			}
			if (element.geometry === "round1Rect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const radius = Math.min(width, height) * (element.presetAdjustment ?? .16667);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 0 L${x(width - radius)} 0 A${x(radius)} ${y(radius)} 0 0 1 1000 ${y(radius)} L1000 1000 H0 Z`;
			}
			if (element.geometry === "round2DiagRect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const radius1 = shortSide * (element.presetAdjustments?.adj1 ?? 16667) / 1e5;
				const radius2 = shortSide * (element.presetAdjustments?.adj2 ?? 0) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(radius1)} 0 L${x(width - radius2)} 0 A${x(radius2)} ${y(radius2)} 0 0 1 1000 ${y(radius2)} L1000 ${y(height - radius1)} A${x(radius1)} ${y(radius1)} 0 0 1 ${x(width - radius1)} 1000 L${x(radius2)} 1000 A${x(radius2)} ${y(radius2)} 0 0 1 0 ${y(height - radius2)} L0 ${y(radius1)} A${x(radius1)} ${y(radius1)} 0 0 1 ${x(radius1)} 0 Z`;
			}
			if (element.geometry === "round2SameRect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const topRadius = shortSide * (element.presetAdjustments?.adj1 ?? 16667) / 1e5;
				const bottomRadius = shortSide * (element.presetAdjustments?.adj2 ?? 0) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(topRadius)} 0 L${x(width - topRadius)} 0 A${x(topRadius)} ${y(topRadius)} 0 0 1 1000 ${y(topRadius)} L1000 ${y(height - bottomRadius)} A${x(bottomRadius)} ${y(bottomRadius)} 0 0 1 ${x(width - bottomRadius)} 1000 L${x(bottomRadius)} 1000 A${x(bottomRadius)} ${y(bottomRadius)} 0 0 1 0 ${y(height - bottomRadius)} L0 ${y(topRadius)} A${x(topRadius)} ${y(topRadius)} 0 0 1 ${x(topRadius)} 0 Z`;
			}
			if (element.geometry === "snipRoundRect") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const radius = shortSide * (element.presetAdjustments?.adj1 ?? 16667) / 1e5;
				const cut = shortSide * (element.presetAdjustments?.adj2 ?? 16667) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(radius)} 0 L${x(width - cut)} 0 L1000 ${y(cut)} L1000 1000 H0 L0 ${y(radius)} A${x(radius)} ${y(radius)} 0 0 1 ${x(radius)} 0 Z`;
			}
			if (element.geometry === "cube") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const inset = Math.min(width, height) * (element.presetAdjustment ?? .25);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 ${y(inset)} L${x(inset)} 0 H1000 V${y(height - inset)} L${x(width - inset)} 1000 H0 Z`;
			}
			if (element.geometry === "frame") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const inset = Math.min(width, height) * Math.max(0, Math.min(5e4, element.presetAdjustments?.adj1 ?? 12500)) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 0 H1000 V1000 H0 Z M${x(inset)} ${y(inset)} V${y(height - inset)} H${x(width - inset)} V${y(inset)} Z`;
			}
			if (element.geometry === "halfFrame") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const maximumAdj2 = 1e5 * width / shortSide;
				const x1 = shortSide * Math.max(0, Math.min(maximumAdj2, element.presetAdjustments?.adj2 ?? 33333)) / 1e5;
				const maximumAdj1 = 1e5 * (height - height * x1 / width) / shortSide;
				const y1 = shortSide * Math.max(0, Math.min(maximumAdj1, element.presetAdjustments?.adj1 ?? 33333)) / 1e5;
				const x2 = width - y1 * width / height;
				const y2 = height - x1 * height / width;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 0 H1000 L${x(x2)} ${y(y1)} L${x(x1)} ${y(y1)} L${x(x1)} ${y(y2)} Z`;
			}
			if (element.geometry === "cornerTabs" || element.geometry === "plaqueTabs" || element.geometry === "squareTabs") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "";
				const inset = Math.hypot(width, height) / 20;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				const rightInset = x(width - inset), bottomInset = y(height - inset), topRight = y(inset), leftRight = x(inset);
				if (element.geometry === "cornerTabs") return `M0 0 L${leftRight} 0 L0 ${topRight} Z M0 ${bottomInset} L${leftRight} 1000 L0 1000 Z M${rightInset} 0 L1000 0 L1000 ${topRight} Z M1000 ${bottomInset} L1000 1000 L${rightInset} 1000 Z`;
				if (element.geometry === "squareTabs") return `M0 0 L${leftRight} 0 L${leftRight} ${topRight} L0 ${topRight} Z M0 ${bottomInset} L${leftRight} ${bottomInset} L${leftRight} 1000 L0 1000 Z M${rightInset} 0 L1000 0 L1000 ${topRight} L${rightInset} ${topRight} Z M${rightInset} ${bottomInset} L1000 ${bottomInset} L1000 1000 L${rightInset} 1000 Z`;
				return `M0 0 L${leftRight} 0 A${leftRight} ${topRight} 0 0 1 0 ${topRight} Z M0 ${bottomInset} A${leftRight} ${topRight} 0 0 1 ${leftRight} 1000 L0 1000 Z M1000 0 L1000 ${topRight} A${leftRight} ${topRight} 0 0 1 ${rightInset} 0 Z M${rightInset} 1000 A${leftRight} ${topRight} 0 0 1 1000 ${bottomInset} L1000 1000 Z`;
			}
			if (element.geometry === "plaque") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const radius = Math.min(width, height) * (element.presetAdjustment ?? .16667);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 ${y(radius)} A${x(radius)} ${y(radius)} 0 0 1 ${x(radius)} 0 L${x(width - radius)} 0 A${x(radius)} ${y(radius)} 0 0 1 1000 ${y(radius)} L1000 ${y(height - radius)} A${x(radius)} ${y(radius)} 0 0 1 ${x(width - radius)} 1000 L${x(radius)} 1000 A${x(radius)} ${y(radius)} 0 0 1 0 ${y(height - radius)} Z`;
			}
			if (element.geometry === "donut") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const inset = Math.min(width, height) * (element.presetAdjustment ?? .25);
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
				const x = (width - Math.min(width, height) * (element.presetAdjustment ?? .5)) / width * 1e3;
				return `M0 0 L${x} 0 L1000 500 L${x} 1000 L0 1000 Z`;
			}
			if (element.geometry === "nonIsoscelesTrapezoid") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const shortSide = Math.min(width, height);
				const leftTop = shortSide * (element.presetAdjustments?.adj1 ?? 25e3) / 1e5;
				const rightTop = width - shortSide * (element.presetAdjustments?.adj2 ?? 25e3) / 1e5;
				const x = (value) => value / width * 1e3;
				return `M0 1000 L${x(leftTop)} 0 L${x(rightTop)} 0 L1000 1000 Z`;
			}
			if (element.geometry === "heart") {
				const innerControlX = 500 + 1e4 / 48, outerControlX = 500 + 49e3 / 48;
				const upperControlY = -1e3 / 3;
				return `M500 250 C${innerControlX} ${upperControlY} ${outerControlX} 250 500 1000 C-520.8333333333335 250 291.66666666666663 ${upperControlY} 500 250 Z`;
			}
			if (element.geometry === "lightningBolt") {
				const scale = 1e3 / 21600;
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
				].map(([pointX, pointY], index) => (index ? "L" : "M") + pointX * scale + " " + pointY * scale).join(" ") + " Z";
			}
			if (element.geometry === "triangle") return `M0 1000 L${(element.presetAdjustment ?? .5) * 1e3} 0 L1000 1000 Z`;
			if (element.geometry === "rtTriangle") return "M0 1000 L0 0 L1000 1000 Z";
			if (element.geometry === "diamond" || element.geometry === "flowChartDecision") return "M0 500 L500 0 L1000 500 L500 1000 Z";
			if (element.geometry === "flowChartOffpageConnector") return "M0 0 H1000 V800 L500 1000 L0 800 Z";
			if (element.geometry === "flowChartDocument") {
				const scale = 1e3 / 21600;
				return `M0 0 H1000 V${17322 * scale} C${10800 * scale} ${17322 * scale} ${10800 * scale} ${23922 * scale} 0 ${20172 * scale} Z`;
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
				const x = 500 + 500 * Math.cos(angle), y = 500 + 500 * Math.sin(angle);
				return `M500 1000 A500 500 0 0 1 0 500 A500 500 0 0 1 500 0 A500 500 0 0 1 1000 500 A500 500 0 0 1 ${x} ${y} L1000 ${y} L1000 1000 Z`;
			}
			if (element.geometry === "flowChartMultidocument") {
				const scale = 1e3 / 21600;
				const unit = (value) => value * scale;
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
				const cut = element.width > 0 ? Math.min(element.width, element.height) * (element.presetAdjustment ?? .25) / element.width * 1e3 : 0;
				return `M0 1000 L${cut} 0 L${1e3 - cut} 0 L1000 1000 Z`;
			}
			if (element.geometry === "octagon") {
				const shortSide = Math.min(element.width, element.height);
				const adjustment = element.presetAdjustment ?? .29289;
				const cutX = element.width > 0 ? shortSide * adjustment / element.width * 1e3 : 0;
				const cutY = element.height > 0 ? shortSide * adjustment / element.height * 1e3 : 0;
				return `M0 ${cutY} L${cutX} 0 H${1e3 - cutX} L1000 ${cutY} V${1e3 - cutY} L${1e3 - cutX} 1000 H${cutX} L0 ${1e3 - cutY} Z`;
			}
			if (element.geometry === "pentagon") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const halfWidth = width * (element.presetAdjustments?.hf ?? 105146) / 2e5;
				const halfHeight = height * (element.presetAdjustments?.vf ?? 110557) / 2e5;
				const dx1 = halfWidth * Math.cos(Math.PI / 10), dx2 = halfWidth * Math.cos(17 * Math.PI / 10);
				const dy1 = halfHeight * Math.sin(Math.PI / 10), dy2 = halfHeight * Math.sin(17 * Math.PI / 10);
				const centerY = height * (element.presetAdjustments?.vf ?? 110557) / 2e5;
				const pointX = (value) => value / width * 1e3, pointY = (value) => value / height * 1e3;
				return `M${pointX(width / 2 - dx1)} ${pointY(centerY - dy1)} L500 0 L${pointX(width / 2 + dx1)} ${pointY(centerY - dy1)} L${pointX(width / 2 + dx2)} ${pointY(centerY - dy2)} L${pointX(width / 2 - dx2)} ${pointY(centerY - dy2)} Z`;
			}
			if (element.geometry === "heptagon") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const halfWidth = width / 2 * (element.presetAdjustments?.hf ?? 102572) / 1e5;
				const halfHeight = height / 2 * (element.presetAdjustments?.vf ?? 105210) / 1e5;
				const dx1 = halfWidth * 97493 / 1e5, dx2 = halfWidth * 78183 / 1e5, dx3 = halfWidth * 43388 / 1e5;
				const y1 = halfHeight - halfHeight * 62349 / 1e5;
				const y2 = halfHeight + halfHeight * 22252 / 1e5;
				const y3 = halfHeight + halfHeight * 90097 / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(width / 2 - dx1)} ${y(y2)} L${x(width / 2 - dx2)} ${y(y1)} L500 0 L${x(width / 2 + dx2)} ${y(y1)} L${x(width / 2 + dx1)} ${y(y2)} L${x(width / 2 + dx3)} ${y(y3)} L${x(width / 2 - dx3)} ${y(y3)} Z`;
			}
			if (element.geometry === "decagon") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const halfHeight = height / 2 * (element.presetAdjustments?.vf ?? 105146) / 1e5;
				const dx1 = width / 2 * Math.cos(Math.PI / 5), dx2 = width / 2 * Math.cos(2 * Math.PI / 5);
				const dy1 = halfHeight * Math.sin(2 * Math.PI / 5), dy2 = halfHeight * Math.sin(Math.PI / 5);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 500 L${x(width / 2 - dx1)} ${y(height / 2 - dy2)} L${x(width / 2 - dx2)} ${y(height / 2 - dy1)} L${x(width / 2 + dx2)} ${y(height / 2 - dy1)} L${x(width / 2 + dx1)} ${y(height / 2 - dy2)} L1000 500 L${x(width / 2 + dx1)} ${y(height / 2 + dy2)} L${x(width / 2 + dx2)} ${y(height / 2 + dy1)} L${x(width / 2 - dx2)} ${y(height / 2 + dy1)} L${x(width / 2 - dx1)} ${y(height / 2 + dy2)} Z`;
			}
			if (element.geometry === "star5") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const horizontalFactor = element.presetAdjustments?.hf ?? 105146, verticalFactor = element.presetAdjustments?.vf ?? 110557;
				const halfWidth = width / 2 * horizontalFactor / 1e5, halfHeight = height / 2 * verticalFactor / 1e5;
				const centerY = height / 2 * verticalFactor / 1e5;
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .19098)) * 2;
				const innerHalfWidth = halfWidth * innerScale, innerHalfHeight = halfHeight * innerScale;
				const dx1 = halfWidth * Math.cos(Math.PI / 10), dx2 = halfWidth * Math.cos(17 * Math.PI / 10);
				const dy1 = halfHeight * Math.sin(Math.PI / 10), dy2 = halfHeight * Math.sin(17 * Math.PI / 10);
				const sdx1 = innerHalfWidth * Math.cos(19 * Math.PI / 10), sdx2 = innerHalfWidth * Math.cos(3 * Math.PI / 10);
				const sdy1 = innerHalfHeight * Math.sin(3 * Math.PI / 10), sdy2 = innerHalfHeight * Math.sin(19 * Math.PI / 10);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(width / 2 - dx1)} ${y(centerY - dy1)} L${x(width / 2 - sdx2)} ${y(centerY - sdy1)} L500 0 L${x(width / 2 + sdx2)} ${y(centerY - sdy1)} L${x(width / 2 + dx1)} ${y(centerY - dy1)} L${x(width / 2 + sdx1)} ${y(centerY - sdy2)} L${x(width / 2 + dx2)} ${y(centerY - dy2)} L500 ${y(centerY + innerHalfHeight)} L${x(width / 2 - dx2)} ${y(centerY - dy2)} L${x(width / 2 - sdx1)} ${y(centerY - sdy2)} Z`;
			}
			if (element.geometry === "star4") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .125)) * 2;
				const innerHalfWidth = width / 2 * innerScale, innerHalfHeight = height / 2 * innerScale;
				const insetX = innerHalfWidth * Math.cos(Math.PI / 4), insetY = innerHalfHeight * Math.sin(Math.PI / 4);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 500 L${x(width / 2 - insetX)} ${y(height / 2 - insetY)} L500 0 L${x(width / 2 + insetX)} ${y(height / 2 - insetY)} L1000 500 L${x(width / 2 + insetX)} ${y(height / 2 + insetY)} L500 1000 L${x(width / 2 - insetX)} ${y(height / 2 + insetY)} Z`;
			}
			if (element.geometry === "star6") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const halfWidth = width / 2 * (element.presetAdjustments?.hf ?? 115470) / 1e5;
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .28868)) * 2;
				const innerHalfWidth = halfWidth * innerScale, innerHalfHeight = height / 2 * innerScale;
				const dx = halfWidth * Math.cos(Math.PI / 6), dy = innerHalfHeight * Math.sin(Math.PI / 3);
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M${x(width / 2 - dx)} ${y(height / 4)} L${x(width / 2 - innerHalfWidth / 2)} ${y(height / 2 - dy)} L500 0 L${x(width / 2 + innerHalfWidth / 2)} ${y(height / 2 - dy)} L${x(width / 2 + dx)} ${y(height / 4)} L${x(width / 2 + innerHalfWidth)} 500 L${x(width / 2 + dx)} ${y(height * 3 / 4)} L${x(width / 2 + innerHalfWidth / 2)} ${y(height / 2 + dy)} L500 1000 L${x(width / 2 - innerHalfWidth / 2)} ${y(height / 2 + dy)} L${x(width / 2 - dx)} ${y(height * 3 / 4)} L${x(width / 2 - innerHalfWidth)} 500 Z`;
			}
			if (element.geometry === "star7") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const horizontalFactor = element.presetAdjustments?.hf ?? 102572, verticalFactor = element.presetAdjustments?.vf ?? 105210;
				const outerHalfWidth = width / 2 * horizontalFactor / 1e5, outerHalfHeight = height / 2 * verticalFactor / 1e5;
				const centerX = width / 2, centerY = height / 2 * verticalFactor / 1e5;
				const outerDx1 = outerHalfWidth * .97493, outerDx2 = outerHalfWidth * .78183, outerDx3 = outerHalfWidth * .43388;
				const outerDy1 = outerHalfHeight * .62349, outerDy2 = outerHalfHeight * .22252, outerDy3 = outerHalfHeight * .90097;
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .34601)) * 2;
				const innerHalfWidth = outerHalfWidth * innerScale, innerHalfHeight = outerHalfHeight * innerScale;
				const innerDx1 = innerHalfWidth * .97493, innerDx2 = innerHalfWidth * .78183, innerDx3 = innerHalfWidth * .43388;
				const innerDy1 = innerHalfHeight * .90097, innerDy2 = innerHalfHeight * .22252, innerDy3 = innerHalfHeight * .62349;
				const x1 = centerX - outerDx1, x2 = centerX - outerDx2, x3 = centerX - outerDx3;
				const x4 = centerX + outerDx3, x5 = centerX + outerDx2, x6 = centerX + outerDx1;
				const y1 = centerY - outerDy1, y2 = centerY + outerDy2, y3 = centerY + outerDy3;
				const sx1 = centerX - innerDx1, sx2 = centerX - innerDx2, sx3 = centerX - innerDx3;
				const sx4 = centerX + innerDx3, sx5 = centerX + innerDx2, sx6 = centerX + innerDx1;
				const sy1 = centerY - innerDy1, sy2 = centerY - innerDy2, sy3 = centerY + innerDy3, sy4 = centerY + innerHalfHeight;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
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
				].map(([pointX, pointY], index) => (index ? "L" : "M") + x(pointX) + " " + y(pointY)).join(" ") + " Z";
			}
			if (element.geometry === "star8") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .375)) * 2;
				const innerHalfWidth = width / 2 * innerScale, innerHalfHeight = height / 2 * innerScale;
				const outerDx = width / 2 * Math.cos(Math.PI / 4), outerDy = height / 2 * Math.sin(Math.PI / 4);
				const innerDx1 = innerHalfWidth * .92388, innerDx2 = innerHalfWidth * .38268;
				const innerDy1 = innerHalfHeight * .92388, innerDy2 = innerHalfHeight * .38268;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				return `M0 500 L${x(width / 2 - innerDx1)} ${y(height / 2 - innerDy2)} L${x(width / 2 - outerDx)} ${y(height / 2 - outerDy)} L${x(width / 2 - innerDx2)} ${y(height / 2 - innerDy1)} L500 0 L${x(width / 2 + innerDx2)} ${y(height / 2 - innerDy1)} L${x(width / 2 + outerDx)} ${y(height / 2 - outerDy)} L${x(width / 2 + innerDx1)} ${y(height / 2 - innerDy2)} L1000 500 L${x(width / 2 + innerDx1)} ${y(height / 2 + innerDy2)} L${x(width / 2 + outerDx)} ${y(height / 2 + outerDy)} L${x(width / 2 + innerDx2)} ${y(height / 2 + innerDy1)} L500 1000 L${x(width / 2 - innerDx2)} ${y(height / 2 + innerDy1)} L${x(width / 2 - outerDx)} ${y(height / 2 + outerDy)} L${x(width / 2 - innerDx1)} ${y(height / 2 + innerDy2)} Z`;
			}
			if (element.geometry === "star10") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const outerHalfWidth = width / 2 * (element.presetAdjustments?.hf ?? 105146) / 1e5;
				const outerDx1 = outerHalfWidth * .95106, outerDx2 = outerHalfWidth * .58779;
				const outerDy1 = height / 2 * .80902, outerDy2 = height / 2 * .30902;
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .42533)) * 2;
				const innerHalfWidth = outerHalfWidth * innerScale, innerHalfHeight = height / 2 * innerScale;
				const innerDx1 = innerHalfWidth * .80902, innerDx2 = innerHalfWidth * .30902;
				const innerDy1 = innerHalfHeight * .95106, innerDy2 = innerHalfHeight * .58779;
				const centerX = width / 2, centerY = height / 2;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
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
				].map(([pointX, pointY], index) => (index ? "L" : "M") + x(pointX) + " " + y(pointY)).join(" ") + " Z";
			}
			if (element.geometry === "star12") {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const centerX = width / 2, centerY = height / 2;
				const outerDx = centerX * Math.cos(Math.PI / 6), outerDy = centerY * Math.sin(Math.PI / 3);
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .375)) * 2;
				const innerHalfWidth = centerX * innerScale, innerHalfHeight = centerY * innerScale;
				const innerDx1 = innerHalfWidth * Math.cos(Math.PI / 12), innerDx2 = innerHalfWidth * Math.cos(Math.PI / 4), innerDx3 = innerHalfWidth * Math.cos(5 * Math.PI / 12);
				const innerDy1 = innerHalfHeight * Math.sin(5 * Math.PI / 12), innerDy2 = innerHalfHeight * Math.sin(Math.PI / 4), innerDy3 = innerHalfHeight * Math.sin(Math.PI / 12);
				const x1 = centerX - outerDx, x3 = width * 3 / 4, x4 = centerX + outerDx;
				const y1 = centerY - outerDy, y3 = height * 3 / 4, y4 = centerY + outerDy;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
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
				].map(([pointX, pointY], index) => (index ? "L" : "M") + x(pointX) + " " + y(pointY)).join(" ") + " Z";
			}
			if ([
				"star16",
				"star24",
				"star32"
			].includes(element.geometry)) {
				const width = element.width, height = element.height;
				if (width <= 0 || height <= 0) return "M0 0 H1000 V1000 H0 Z";
				const pointCount = Number(element.geometry.slice(4));
				const innerScale = Math.max(0, Math.min(.5, element.presetAdjustment ?? .375)) * 2;
				const centerX = width / 2, centerY = height / 2;
				const guideFactor = (value) => Math.round(value * 1e5) / 1e5;
				const x = (value) => value / width * 1e3, y = (value) => value / height * 1e3;
				const points = [];
				for (let index = 0; index < pointCount * 2; index++) {
					const angle = Math.PI + index * Math.PI / pointCount;
					const radius = index % 2 === 0 ? 1 : innerScale;
					const usesExactTrig = element.geometry === "star24" && index % 2 === 0 || element.geometry === "star32" && index % 8 === 0;
					const factor = (value) => usesExactTrig ? value : guideFactor(value);
					points.push([centerX + centerX * radius * factor(Math.cos(angle)), centerY + centerY * radius * factor(Math.sin(angle))]);
				}
				return points.map(([pointX, pointY], index) => (index ? "L" : "M") + x(pointX) + " " + y(pointY)).join(" ") + " Z";
			}
			if (element.geometry === "dodecagon") {
				const scale = 1e3 / 21600;
				const x1 = 2894 * scale, x2 = 7906 * scale, x3 = 13694 * scale, x4 = 18706 * scale;
				return `M0 ${x2} L${x1} ${x1} L${x2} 0 H${x3} L${x4} ${x1} L1000 ${x2} V${x3} L${x4} ${x4} L${x3} 1000 H${x2} L${x1} ${x4} L0 ${x3} Z`;
			}
			if (element.geometry === "parallelogram" || element.geometry === "chevron" || element.geometry === "hexagon") {
				const adjustment = element.presetAdjustment ?? (element.geometry === "parallelogram" ? .25 : element.geometry === "hexagon" ? .25 : .5);
				const shortSide = Math.min(element.width, element.height);
				const cut = shortSide > 0 && element.width > 0 ? shortSide * adjustment / element.width * 1e3 : 0;
				if (element.geometry === "parallelogram") return `M0 1000 L${cut} 0 H1000 L${1e3 - cut} 1000 Z`;
				if (element.geometry === "chevron") return `M0 0 H${1e3 - cut} L1000 500 L${1e3 - cut} 1000 H0 L${cut} 500 Z`;
				return `M0 500 L${cut} ${.211325 * 1e3} H${1e3 - cut} L1000 500 L${1e3 - cut} 788.675 H${cut} Z`;
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
			const brightness = customPathBrightness(path.fillMode);
			if (!animatedFill && isPatternFillCss(element.fill)) return `url(#${customGeometryPatternId(element, path.fillMode)})`;
			if (!animatedFill && /^(?:linear|radial)-gradient\(/i.test(element.fill)) return `url(#${customGradientId(element, "fill", path.fillMode)})`;
			return applyColorOpacity(animatedFill || element.fill, animationOpacity(element, "fill.opacity"), brightness);
		}
		function customGeometryFillOpacity(element, path) {
			if (!path.fill || !animationOnOff(element, "fill.on") || animationColor(element, "fillcolor")) return void 0;
			return isPatternFillCss(element.fill) ? animationOpacity(element, "fill.opacity") : void 0;
		}
		function customPathBrightness(mode) {
			return mode ? {
				lighten: .4,
				lightenLess: .2,
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
			if (strokeWidth > 0 && !animatedStroke && element.strokePattern) return `url(#${customGeometryStrokePatternId(element)})`;
			if (path.stroke && strokeWidth > 0 && !animatedStroke && /^(?:linear|radial)-gradient\(/i.test(element.stroke)) return `url(#${customGradientId(element, "stroke")})`;
			const stroke = applyColorOpacity(animatedStroke || element.stroke, animationOpacity(element, "stroke.opacity"));
			return path.stroke && strokeWidth > 0 && stroke !== "transparent" && !stroke.includes("gradient") ? stroke : "none";
		}
		function customGradientId(element, kind, mode) {
			return `${customGradientScope}-${kind}-${element.id.replace(/[^a-zA-Z0-9_-]/g, "_")}${mode ? `-${mode}` : ""}`;
		}
		function customGeometryPatternModes(element) {
			if (!isPatternFillCss(element.fill)) return [];
			return [.../* @__PURE__ */ new Set([void 0, ...element.customPaths?.flatMap((path) => path.fillMode ? [path.fillMode] : []) || []])];
		}
		function customGeometryPatternId(element, mode) {
			return `${customGradientScope}-pattern-${element.id.replace(/[^a-zA-Z0-9_-]/g, "_")}${mode ? `-${mode}` : ""}`;
		}
		function customGeometryStrokePatternId(element) {
			return `${customGradientScope}-stroke-pattern-${element.id.replace(/[^a-zA-Z0-9_-]/g, "_")}`;
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
				const focusColor = focusEllipse ? splitGradientStops(focusEllipse[1])[0]?.replace(/\s+[-+\d.]+%$/, "").trim() : focusRect ? splitGradientStops(focusRect[1])[0]?.trim() : void 0;
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
			const modifiers = [...new Set(element.customPaths?.map((path) => path.fillMode).filter((mode) => mode))];
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
			return gradient?.kind === "radial" ? gradient.focus : void 0;
		}
		function customGeometryStrokeWidth(element) {
			return `${animationStrokeWidth(element) * scale.value}px`;
		}
		function customGeometryStrokeStyle(element) {
			const strokeDash = animationStrokeDash(element);
			const strokeWidth = animationStrokeWidth(element) * scale.value;
			const cap = element.strokeCap || "butt";
			const pattern = svgStrokeDashPattern(strokeDash, cap);
			return {
				...pattern ? { strokeDasharray: pattern.map((length) => `${length * strokeWidth}px`).join(" ") } : {},
				strokeOpacity: !animationColor(element, "stroke.color") && (/gradient\(/i.test(element.stroke) || Boolean(element.strokePattern)) ? animationOpacity(element, "stroke.opacity") : void 0,
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
			const animatedStroke = animationColor(element, "stroke.color");
			if (!animatedStroke && element.strokePattern) return `url(#${customGeometryStrokePatternId(element)})`;
			return applyColorOpacity(animatedStroke || (element.stroke === "transparent" ? "#222" : element.stroke), animationOpacity(element, "stroke.opacity"));
		}
		function lineRenderStyle(element) {
			const stroke = animationColor(element, "stroke.color") || (element.stroke === "transparent" ? "#222" : element.stroke);
			return {
				width: `${Math.max(element.width * scale.value, 1)}px`,
				height: `${Math.max(element.height * scale.value, 1)}px`,
				color: applyColorOpacity(stroke, animationOpacity(element, "stroke.opacity"))
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
				sm: .7,
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
			const property = side === "head" ? "stroke.startArrow" : "stroke.endArrow";
			const baseline = side === "head" ? element.lineHeadEnd : element.lineTailEnd;
			let type = baseline?.type || "none";
			for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "lineArrow" && item.lineArrowProperty === property) || []) {
				const progress = animationProgress(action);
				if (progress === -2) {
					type = baseline?.type || "none";
					continue;
				}
				if (progress === null || progress < 0) continue;
				if (action.lineArrowKeyframes?.length) for (const frame of action.lineArrowKeyframes) {
					if (frame.offset > progress) break;
					type = lineArrowType(frame.value);
				}
				else if (action.lineArrowFrom !== void 0 && progress < 1) type = lineArrowType(action.lineArrowFrom);
				else if (progress >= 1) {
					const value = action.lineArrowTo ?? action.lineArrowFrom;
					if (value !== void 0) type = lineArrowType(value);
				}
			}
			const baselineWidth = baseline?.width || "med";
			const baselineLength = baseline?.length || "med";
			const animatedSize = (property, initial) => {
				let value = initial;
				for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "lineArrowSize" && item.lineArrowSizeProperty === property) || []) {
					const progress = animationProgress(action);
					if (progress === -2) {
						value = initial;
						continue;
					}
					if (progress === null || progress < 0) continue;
					if (action.lineArrowSizeKeyframes?.length) for (const frame of action.lineArrowSizeKeyframes) {
						if (frame.offset > progress) break;
						value = lineArrowSize(frame.value);
					}
					else if (action.lineArrowSizeFrom !== void 0 && progress < 1) value = lineArrowSize(action.lineArrowSizeFrom);
					else if (progress >= 1) {
						const finalValue = action.lineArrowSizeTo ?? action.lineArrowSizeFrom;
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
			const outerShadows = (element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []).filter((shadow) => !shadow.inset);
			const innerShadowFilter = customGeometryInnerShadowFilterReference(element);
			if (!outerShadows.length && !innerShadowFilter || !animationOnOff(element, "shadow.on")) return {};
			return { filter: [innerShadowFilter, ...outerShadows.map((shadow) => `drop-shadow(${shadowCss(shadow, applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity")), .5)})`)].filter((filter) => Boolean(filter)).join(" ") };
		}
		function customGeometryInnerShadows(element) {
			return (element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []).filter((shadow) => shadow.inset);
		}
		function customGeometryInnerShadowFilterId(element) {
			return `${customGradientScope}-inner-shadow-${element.id}`.replace(/[^a-zA-Z0-9_-]/g, "_");
		}
		function customGeometryInnerShadowFilterReference(element) {
			return customGeometryInnerShadows(element).length && animationOnOff(element, "shadow.on") ? `url(#${customGeometryInnerShadowFilterId(element)})` : void 0;
		}
		function customGeometryInnerShadowFilterRegion(element) {
			let paddingX = 0;
			let paddingY = 0;
			for (const shadow of customGeometryInnerShadows(element)) {
				const angle = (shadow.angle || 0) * Math.PI / 180;
				const blur = Math.max(0, shadow.blur) * 1.5;
				paddingX = Math.max(paddingX, (Math.abs(Math.cos(angle) * shadow.distance) + blur) / Math.max(1, element.width));
				paddingY = Math.max(paddingY, (Math.abs(Math.sin(angle) * shadow.distance) + blur) / Math.max(1, element.height));
			}
			return {
				x: -paddingX,
				y: -paddingY,
				width: 1 + paddingX * 2,
				height: 1 + paddingY * 2
			};
		}
		function customGeometryInnerShadowDeviation(element, shadow) {
			const blur = Math.max(0, shadow.blur) / 2;
			return `${blur / Math.max(1, element.width)} ${blur / Math.max(1, element.height)}`;
		}
		function customGeometryInnerShadowOffset(element, shadow) {
			const angle = (shadow.angle || 0) * Math.PI / 180;
			return {
				dx: Math.cos(angle) * shadow.distance / Math.max(1, element.width),
				dy: Math.sin(angle) * shadow.distance / Math.max(1, element.height)
			};
		}
		function customGeometryInnerShadowColor(element, shadow) {
			return applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity"));
		}
		function shadowCss(shadow, color, blurScale = 1) {
			const angle = (shadow.angle || 0) * Math.PI / 180;
			return [
				...shadow.inset ? ["inset"] : [],
				Math.cos(angle) * shadow.distance * scale.value + "px",
				Math.sin(angle) * shadow.distance * scale.value + "px",
				shadow.blur * blurScale * scale.value + "px",
				color
			].join(" ");
		}
		function pictureInnerShadowStyle(element) {
			const innerShadows = (element.shadows?.length ? element.shadows : element.shadow ? [element.shadow] : []).filter((shadow) => shadow.inset);
			if (!innerShadows.length || !animationOnOff(element, "shadow.on")) return void 0;
			return { boxShadow: innerShadows.map((shadow) => shadowCss(shadow, applyColorOpacity(animationColor(element, "shadow.color") || shadow.color, animationOpacity(element, "shadow.opacity")))).join(", ") };
		}
		function imageStyle(element) {
			const hasCropAnimation = animationsByTarget.value.get(element.id)?.some((item) => item.effect === "imageCrop");
			const crop = element.imageCrop || hasCropAnimation ? imageCrop(element) : void 0;
			const grayscale = element.kind === "picture" && imageGrayscale(element);
			const alphaInverse = element.kind === "picture" && element.imageAlphaInverse !== void 0 ? `url(#${imageAlphaInverseFilterId(element)})` : void 0;
			const effectDagAlphaInverse = element.kind === "picture" && element.imageEffectDagAlphaInverse !== void 0 ? `url(#${imageAlphaInverseFilterId(element, true)})` : void 0;
			const grayscaleFilter = grayscale ? "grayscale(1)" : void 0;
			const imageEffects = element.imageAlphaInverseAfterGrayscale ? [grayscaleFilter, alphaInverse] : [alphaInverse, grayscaleFilter];
			const filters = [
				imageToneEnabled(element) ? `url(#${imageToneFilterId(element)})` : void 0,
				...imageEffects,
				effectDagAlphaInverse
			].filter((filter) => Boolean(filter));
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
			const base = element.imageCrop || {
				left: 0,
				top: 0,
				right: 0,
				bottom: 0
			};
			const crop = { ...base };
			for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "imageCrop") || []) {
				const property = action.imageCropProperty;
				if (!property || action.imageCropTo === void 0 && !action.imageCropKeyframes?.length) continue;
				const progress = animationProgress(action);
				if (progress === -2) {
					crop[property] = base[property];
					continue;
				}
				if (progress === null || progress < 0) continue;
				const value = action.imageCropKeyframes?.length ? animationKeyframeValue(action.imageCropKeyframes, action.imageCropKeyframeMode, progress) : (action.imageCropFrom ?? crop[property]) + ((action.imageCropTo ?? 0) - (action.imageCropFrom ?? crop[property])) * progress;
				if (value >= 0 && value < 1 && value + crop[property === "left" ? "right" : property === "right" ? "left" : property === "top" ? "bottom" : "top"] < 1) crop[property] = value;
			}
			return crop;
		}
		function imageGrayscale(element) {
			let grayscale = element.imageGrayscale === true;
			for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "grayscale") || []) {
				const progress = animationProgress(action);
				if (progress === null || progress < 0) continue;
				let value = action.grayscaleFrom ?? grayscale;
				if (action.grayscaleKeyframes?.length) for (const frame of action.grayscaleKeyframes) {
					if (frame.offset > progress) break;
					value = frame.value;
				}
				else if (progress >= 1) value = action.grayscaleTo ?? value;
				grayscale = value;
			}
			return grayscale;
		}
		function imageTone(element) {
			const tone = {
				gain: 1,
				blacklevel: 0,
				gamma: 1
			};
			for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "imageTone") || []) {
				const property = action.imageToneProperty;
				if (!property || action.imageToneTo === void 0 && !action.imageToneKeyframes?.length) continue;
				const progress = animationProgress(action);
				if (progress === -2) {
					tone[property] = property === "blacklevel" ? 0 : 1;
					continue;
				}
				if (progress === null || progress < 0) continue;
				const from = action.imageToneFrom ?? tone[property];
				const value = action.imageToneKeyframes?.length ? animationKeyframeValue(action.imageToneKeyframes, action.imageToneKeyframeMode, progress) : from + ((action.imageToneTo ?? from) - from) * progress;
				if (Number.isFinite(value)) tone[property] = value;
			}
			return tone;
		}
		function imageToneEnabled(element) {
			return element.kind === "picture" && Boolean(animationsByTarget.value.get(element.id)?.some((item) => item.effect === "imageTone"));
		}
		function imageToneFilterId(element) {
			return `${customGradientScope}-image-tone-${element.id}`.replace(/[^a-zA-Z0-9_-]/g, "_");
		}
		function imageAlphaInverseFilterId(element, effectDag = false) {
			return (customGradientScope + "-image-alpha-inverse-" + element.id + (effectDag ? "-effect-dag" : "")).replace(/[^a-zA-Z0-9_-]/g, "_");
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
		function finishTimedMedia(element, refresh = true, targetEnded = false) {
			const endedAt = performance.now();
			let changed = false;
			const media = Array.from(stage.value?.querySelectorAll("[data-media-id]") || []).find((item) => item.dataset.mediaId === element.id);
			if (media) {
				delete media.dataset.mediaActionEndAt;
				mediaBookmarkPlaybackStarted.delete(media);
				lastMediaBookmarkTimes.delete(media);
			}
			timelineActions.value.forEach((action, index) => {
				if (action.effect !== "media" || action.targetId !== element.id || !action.mediaActionKey || !executedMediaActions.has(index)) return;
				completedMediaActions.set(action.mediaActionKey, endedAt);
				if (targetEnded) endedMediaTargets.set(action.mediaActionKey, endedAt);
				const timer = mediaActionEndTimers.get(action.mediaActionKey);
				if (timer) clearTimeout(timer);
				mediaActionEndTimers.delete(action.mediaActionKey);
				changed = true;
			});
			if (changed && refresh) refreshTimeline();
		}
		function runTimedMediaActions() {
			timelineActions.value.forEach((action, index) => {
				if (action.effect !== "media" || !action.mediaCommand || executedMediaActions.has(index)) return;
				const startedAt = timelineStart(action);
				const stoppedAt = timelineEnd(action);
				if (stoppedAt !== void 0 && startedAt !== void 0 && startedAt + action.delayMs >= stoppedAt) return;
				if (startedAt === void 0 || startedAt === Number.NEGATIVE_INFINITY || animationClock.value < startedAt + action.delayMs) return;
				executedMediaActions.add(index);
				if (action.mediaActionKey) scheduledMediaActions.set(action.mediaActionKey, startedAt + action.delayMs);
				if (action.mediaCommand === "stopAudio") {
					emit("stop-all-audio");
					for (const element of props.slide.elements) {
						if (element.kind !== "audio") continue;
						const audio = Array.from(stage.value?.querySelectorAll("[data-media-id]") || []).find((item) => item.dataset.mediaId === element.id);
						if (!audio) continue;
						if (audio.readyState > 0) audio.currentTime = (element.mediaTrim?.startMs || 0) / 1e3;
						finishTimedMedia(element, false);
					}
					runSlideTriggeredAnimations("onStopAudio");
					refreshTimeline();
					return;
				}
				const element = props.slide.elements.find((item) => item.id === action.targetId);
				const media = Array.from(stage.value?.querySelectorAll("[data-media-id]") || []).find((item) => item.dataset.mediaId === action.targetId);
				if (!element || !media) return;
				const startsPlayback = action.mediaCommand === "play" || action.mediaCommand === "togglePause" && media.paused;
				if (action.mediaActionKey && (!startsPlayback || !media.paused)) startedMediaActions.set(action.mediaActionKey, performance.now());
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
					media.play().then(() => {
						if (action.mediaActionKey && !startedMediaActions.has(action.mediaActionKey)) {
							startedMediaActions.set(action.mediaActionKey, performance.now());
							refreshTimeline();
						}
					}).catch((error) => {
						if (error instanceof DOMException && error.name === "NotAllowedError" && !blockedMediaPlaybackIds.value.includes(element.id)) blockedMediaPlaybackIds.value = [...blockedMediaPlaybackIds.value, element.id];
					});
				};
				if (action.mediaCommand === "play") {
					if (action.mediaStartSeconds !== void 0) seek(action.mediaStartSeconds);
					else if (media.currentTime < (element.mediaTrim?.startMs || 0) / 1e3) seek((element.mediaTrim?.startMs || 0) / 1e3);
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
					seek((element.mediaTrim?.startMs || 0) / 1e3);
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
				if (sibling.sequenceKey !== action.sequenceKey || (sibling.sequenceOrder ?? -1) <= (action.sequenceOrder ?? -1)) return false;
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
			if (elapsedWallMs >= warp.groupPlaybackDurationMs && (cyclePositionMs < .01 || cycleDurationMs - cyclePositionMs < .01)) cyclePositionMs = cycleDurationMs;
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
			const start = points?.[0];
			const end = points?.[points.length - 1];
			if (!start || !end) return [0, 0];
			const repeats = completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1);
			return [(end.x - start.x) * repeats, (end.y - start.y) * repeats];
		}
		function animationScale(element) {
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => ["scale", "zoom"].includes(action.effect) && !action.paragraphRange && !action.characterRange && !action.iteration) || [];
			let value = [1, 1];
			for (const action of actions) {
				const progress = animationProgress(action);
				if (progress === -2) {
					if (action.scaleAxis === "x") value = [1, value[1]];
					else if (action.scaleAxis === "y") value = [value[0], 1];
					else value = [1, 1];
					continue;
				}
				if (progress === null || progress < 0) {
					if (action.effect === "zoom" && action.direction === "in" && action.scaleFrom) value = action.scaleFrom;
					continue;
				}
				if (!action.scaleFrom || !action.scaleTo) continue;
				const animatedScale = animationScaleValue(action, progress);
				if (action.scaleAxis === "x") value = [animatedScale[0], value[1]];
				else if (action.scaleAxis === "y") value = [value[0], animatedScale[1]];
				else value = animatedScale;
			}
			return value;
		}
		function animationScaleValue(action, progress, iterationIndex) {
			const [fromX, fromY] = action.scaleFrom;
			const [toX, toY] = action.scaleTo;
			const completedRepeats = action.scaleAccumulate ? completedAnimationRepeats(action, iterationIndex) * (action.reversePlayback ? -1 : 1) : 0;
			const valueAt = (axis, from, to) => {
				const frames = action.scaleKeyframes;
				if (!frames?.length || action.scaleAxis !== axis) return from + (to - from) * progress;
				let low = 0;
				let high = frames.length;
				while (low < high) {
					const middle = low + high >>> 1;
					if (frames[middle].offset < progress) low = middle + 1;
					else high = middle;
				}
				if (low === 0) return frames[0].value;
				if (low === frames.length) return frames.at(-1).value;
				const before = frames[low - 1];
				const after = frames[low];
				if (action.scaleKeyframeMode === "discrete") return progress === after.offset ? after.value : before.value;
				const ratio = (progress - before.offset) / (after.offset - before.offset);
				return before.value + (after.value - before.value) * ratio;
			};
			return [valueAt("x", fromX, toX) + (toX - fromX) * completedRepeats, valueAt("y", fromY, toY) + (toY - fromY) * completedRepeats];
		}
		function animationRotation(element) {
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "rotation" && !action.paragraphRange && !action.characterRange && !action.iteration) || [];
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
				const rotation = (action.rotationKeyframes?.length ? animationKeyframeValue(action.rotationKeyframes, action.rotationKeyframeMode, progress) : action.rotationFrom + delta * progress) + repeatOffset;
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "opacity" || action.opacityProperty !== property) return false;
				const hasTextRange = Boolean(action.paragraphRange || action.characterRange || action.iteration);
				if (paragraphIndex === void 0) return !hasTextRange;
				const iterationIndex = characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd);
				return hasTextRange && (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart !== void 0 && characterEnd !== void 0 && characterStart < action.characterRange.end && characterEnd > action.characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
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
				if (formulaSamples?.length) opacity = animationKeyframeValue(formulaSamples, "lin", progress);
				else if (keyframes?.length) opacity = animationKeyframeValue(keyframes, action.opacityKeyframeMode, progress);
				else opacity = action.opacityFrom + (action.opacityTo - action.opacityFrom) * progress;
			}
			return opacity;
		}
		function animationOnOff(element, property) {
			let enabled = true;
			for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "onOff" && item.onOffProperty === property) || []) {
				const progress = animationProgress(action);
				if (progress === -2) {
					enabled = true;
					continue;
				}
				if (progress === null || progress < 0) continue;
				if (action.onOffKeyframes?.length) for (const frame of action.onOffKeyframes) {
					if (frame.offset > progress) break;
					enabled = frame.value;
				}
				else if (action.onOffFrom !== void 0 && progress < 1) enabled = action.onOffFrom;
				else if (progress >= 1) enabled = action.onOffTo ?? action.onOffFrom ?? enabled;
			}
			return enabled;
		}
		function animationStrokeDash(element) {
			let strokeDash = element.strokeDash;
			for (const action of animationsByTarget.value.get(element.id)?.filter((item) => item.effect === "strokeDash") || []) {
				const progress = animationProgress(action);
				if (progress === -2) {
					strokeDash = element.strokeDash;
					continue;
				}
				if (progress === null || progress < 0) continue;
				if (action.strokeDashKeyframes?.length) for (const frame of action.strokeDashKeyframes) {
					if (frame.offset > progress) break;
					strokeDash = frame.value;
				}
				else if (action.strokeDashFrom !== void 0 && progress < 1) strokeDash = action.strokeDashFrom;
				else if (progress >= 1) strokeDash = action.strokeDashTo ?? action.strokeDashFrom ?? strokeDash;
			}
			return strokeDash;
		}
		function animationStrokeWidth(element) {
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "strokeWidth") || [];
			let strokeWidth = element.strokeWidth;
			for (const action of actions) {
				const progress = animationProgress(action);
				if (progress === -2) {
					strokeWidth = element.strokeWidth;
					continue;
				}
				if (progress === null || progress < 0 || action.strokeWidthFrom === void 0 || action.strokeWidthTo === void 0) continue;
				strokeWidth = action.strokeWidthFormulaSamples?.length ? animationKeyframeValue(action.strokeWidthFormulaSamples, "lin", progress) : action.strokeWidthKeyframes?.length ? animationKeyframeValue(action.strokeWidthKeyframes, action.strokeWidthKeyframeMode, progress) : action.strokeWidthFrom + (action.strokeWidthTo - action.strokeWidthFrom) * progress;
			}
			if (strokeWidth === 0 && (animationColor(element, "stroke.color") || element.stroke) !== "transparent" && scale.value > 0) return 1 / (scale.value * pixelRatio.value);
			return strokeWidth;
		}
		function animationColor(element, property, paragraphIndex, characterStart, characterEnd) {
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "color" || action.colorProperty !== property) return false;
				const paragraphRange = action.paragraphRange;
				const characterRange = action.characterRange;
				const iterationIndex = characterStart === void 0 || characterEnd === void 0 ? void 0 : iterationIndexFor(action, characterStart, characterEnd);
				return (!paragraphRange || paragraphIndex !== void 0 && paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart !== void 0 && characterEnd !== void 0 && characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
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
		function runHasTextFill(element, run, paragraphIndex, characterStart, characterEnd) {
			return Boolean(run.fillPattern || run.fillGradient) && !animationColor(element, "style.color", paragraphIndex, characterStart, characterEnd);
		}
		function textRangeOpacityAnimationStyle(element, paragraphIndex, characterStart, characterEnd) {
			const opacity = animationOpacity(element, "style.opacity", paragraphIndex, characterStart, characterEnd);
			return opacity === 1 ? {} : { opacity };
		}
		function textRunAnimationStyle(element, paragraphIndex, characterStart, characterEnd, run, fontScale = element.textFontScale ?? 1, script = "latin", separateDecorations = false, filterKey = `${paragraphIndex}-${characterStart}-${characterEnd}`) {
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "fontSize") return false;
				const paragraphRange = action.paragraphRange;
				const characterRange = action.characterRange;
				const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
				return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
			let factor = 1;
			for (const action of actions) {
				const progress = animationProgress(action, iterationIndexFor(action, characterStart, characterEnd));
				if (progress === -2) {
					factor = 1;
					continue;
				}
				if (progress === null || progress < 0 || action.fontSizeFrom === void 0 || action.fontSizeTo === void 0) continue;
				factor = action.fontSizeFormulaSamples?.length ? animationKeyframeValue(action.fontSizeFormulaSamples, "lin", progress) : action.fontSizeKeyframes?.length ? animationKeyframeValue(action.fontSizeKeyframes, action.fontSizeKeyframeMode, progress) : action.fontSizeFrom + (action.fontSizeTo - action.fontSizeFrom) * progress;
			}
			const baseFontWeight = run.bold ? 700 : 400;
			let weight;
			const weightActions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "fontWeight") return false;
				const paragraphRange = action.paragraphRange;
				const characterRange = action.characterRange;
				const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
				return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
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
				weight = action.fontWeightFormulaSamples?.length ? animationKeyframeValue(action.fontWeightFormulaSamples, "lin", progress) : action.fontWeightKeyframes?.length ? animationKeyframeValue(action.fontWeightKeyframes, action.fontWeightKeyframeMode, progress) : action.fontWeightFrom + (action.fontWeightTo - action.fontWeightFrom) * progress;
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
			const textStyleActions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "textStyle") return false;
				const paragraphRange = action.paragraphRange;
				const characterRange = action.characterRange;
				const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
				return (!paragraphRange || paragraphIndex >= paragraphRange.start && paragraphIndex <= paragraphRange.end) && (!characterRange || characterStart < characterRange.end && characterEnd > characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
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
				else if (action.textStyleKeyframes?.length) {
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
			const embossOffset = Math.max(.5, (run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value / 12);
			const textDecorationLine = [underline ?? run.underline ? "underline" : void 0, lineThrough ?? run.lineThrough ? "line-through" : void 0].filter(Boolean).join(" ");
			const textDecorationStyles = separateDecorations ? {
				...underline === void 0 ? {} : { "--pptx-underline-line": underline ? "underline" : "none" },
				...lineThrough === void 0 ? {} : { textDecorationLine: lineThrough ? "line-through" : "none" }
			} : underline === void 0 && lineThrough === void 0 ? {} : { textDecorationLine: textDecorationLine || "none" };
			const filter = [!textShadowChanged || textShadow !== "none" ? textInnerShadowFilterReference(element, run, paragraphIndex, characterStart, characterEnd, filterKey) : void 0, textEmbossChanged && textEmbossEnabled ? `drop-shadow(-${embossOffset}px -${embossOffset}px 0 rgba(255, 255, 255, 0.65)) drop-shadow(${embossOffset}px ${embossOffset}px 0 rgba(0, 0, 0, 0.65))` : void 0].filter(Boolean).join(" ");
			return {
				...factor === 1 ? {} : { fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value * factor}px` },
				...weight === void 0 ? {} : { fontWeight: weight },
				...fontFamily === void 0 ? {} : { fontFamily },
				...italic === void 0 ? {} : { fontStyle: italic ? "italic" : "normal" },
				...textDecorationStyles,
				...verticalAlignChanged ? { verticalAlign } : {},
				...textOutlineChanged ? { WebkitTextStroke: textOutline } : {},
				...textShadowChanged ? { textShadow } : {},
				...textShadowChanged || textEmbossChanged ? { filter: filter || "none" } : {}
			};
		}
		function textTransformAnimationStyle(element, paragraphIndex, characterStart, characterEnd) {
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				const iterationIndex = iterationIndexFor(action, characterStart, characterEnd);
				if (![
					"scale",
					"rotation",
					"motion",
					"slide"
				].includes(action.effect) || !(action.paragraphRange || action.characterRange || action.iteration)) return false;
				return (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start) && (!action.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
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
					const angle = (action.rotationKeyframes?.length ? animationKeyframeValue(action.rotationKeyframes, action.rotationKeyframeMode, progress) : action.rotationFrom + delta * progress) + repeatOffset;
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
			const textRange = paragraphIndex !== void 0 && characterStart !== void 0 && characterEnd !== void 0;
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "motion" && (textRange ? Boolean(action.paragraphRange || action.characterRange || action.iteration) && (!action.paragraphRange || paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end) && (!action.characterRange || characterStart < action.characterRange.end && characterEnd > action.characterRange.start) : !action.paragraphRange && !action.characterRange && !action.iteration)) || [];
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
				if (progress !== null && progress >= 0 && samples?.length) {
					const samplePosition = Math.min(samples.length - 1, progress * (samples.length - 1));
					const from = samples[Math.floor(samplePosition)];
					const to = samples[Math.ceil(samplePosition)];
					const ratio = samplePosition - Math.floor(samplePosition);
					const offset = motionAccumulationOffset(action, iterationIndex);
					const motion = [from.x + (to.x - from.x) * ratio + offset[0], from.y + (to.y - from.y) * ratio + offset[1]];
					if (action.motionCoordinateProperty === "ppt_x") value = [motion[0], value[1]];
					else if (action.motionCoordinateProperty === "ppt_y") value = [value[0], motion[1]];
					else value = motion;
					continue;
				}
				const path = action.motionPath;
				if (progress === null || progress < 0 || !path?.length) continue;
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
			let value = [0, 0];
			const textRange = paragraphIndex !== void 0 || characterStart !== void 0;
			const actions = animationsByTarget.value.get(element.id)?.filter((item) => {
				if (item.effect !== "slide") return false;
				const iterationIndex = characterStart !== void 0 && characterEnd !== void 0 ? iterationIndexFor(item, characterStart, characterEnd) : void 0;
				if (!textRange) return !item.paragraphRange && !item.characterRange && !item.iteration;
				return (Boolean(item.paragraphRange || item.characterRange) || iterationIndex !== void 0 && iterationIndex >= 0) && (!item.paragraphRange || paragraphIndex !== void 0 && paragraphIndex >= item.paragraphRange.start && paragraphIndex <= item.paragraphRange.end) && (!item.characterRange || characterStart !== void 0 && characterEnd !== void 0 && characterStart < item.characterRange.end && characterEnd > item.characterRange.start) && (!item.iteration || iterationIndex !== void 0 && iterationIndex >= 0);
			}) || [];
			for (const action of actions) {
				const progress = animationProgress(action, characterStart !== void 0 && characterEnd !== void 0 ? iterationIndexFor(action, characterStart, characterEnd) : void 0);
				if (progress === -2) {
					value = [0, 0];
					continue;
				}
				if (progress === null) continue;
				const phase = progress < 0 ? 0 : progress;
				const distance = action.direction === "in" ? 1 - phase : phase;
				value = action.slideFrom === "left" ? [-props.width * distance, 0] : action.slideFrom === "right" ? [props.width * distance, 0] : action.slideFrom === "top" ? [0, -props.height * distance] : action.slideFrom === "bottom" ? [0, props.height * distance] : action.slideFrom === "topLeft" ? [-props.width * distance, -props.height * distance] : action.slideFrom === "topRight" ? [props.width * distance, -props.height * distance] : action.slideFrom === "bottomLeft" ? [-props.width * distance, props.height * distance] : action.slideFrom === "bottomRight" ? [props.width * distance, props.height * distance] : [0, 0];
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "wipe") return false;
				if (characterStart !== void 0 && characterEnd !== void 0) {
					const overlaps = Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
					const iterationIndex = action.iteration ? iterationIndexFor(action, characterStart, characterEnd) : void 0;
					return overlaps || iterationIndex !== void 0 && iterationIndex >= 0;
				}
				return paragraphIndex === void 0 ? !action.paragraphRange && !action.characterRange : Boolean(action.paragraphRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
			}) || [];
			const first = actions[0];
			if (!first?.wipeDirection) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "blinds") return false;
				if (paragraphIndex !== void 0) return Boolean(!action.iteration && action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
				if (characterStart !== void 0 && characterEnd !== void 0) {
					const overlaps = Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
					const iteratesCharacter = Boolean(action.iteration && iterationIndexFor(action, characterStart, characterEnd) !== void 0);
					return overlaps || iteratesCharacter;
				}
				return !action.characterRange && !action.paragraphRange;
			}) || [];
			const first = actions[0];
			if (!first?.blindsOrientation) return {};
			let progress = first.direction === "in" ? 0 : 1;
			let orientation = first.blindsOrientation;
			let reset = false;
			for (const action of actions) {
				const phase = animationProgress(action, animationIterationIndex(action, paragraphIndex, characterStart, characterEnd));
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				if (action.effect !== "checker") return false;
				if (paragraphIndex !== void 0) return Boolean(action.paragraphRange && !action.characterRange && paragraphIndex >= action.paragraphRange.start && paragraphIndex <= action.paragraphRange.end);
				if (characterStart !== void 0 && characterEnd !== void 0) return Boolean(action.characterRange && !action.paragraphRange && characterStart < action.characterRange.end && characterEnd > action.characterRange.start);
				return !action.paragraphRange && !action.characterRange;
			}) || [];
			const first = actions[0];
			if (!first?.checkerOrientation) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "shape" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || [];
			const first = actions[0];
			if (!first?.shapeFilter || !first.shapeDirection) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "dissolve" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || [];
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "wheel" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || [];
			const first = actions[0];
			if (!first?.wheelSpokes) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "randomBars" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || [];
			const first = actions[0];
			if (!first?.randomBarOrientation) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "strips" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || [];
			const first = actions[0];
			if (!first?.stripsDirection) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => action.effect === "barn" && animationTargetsRange(action, paragraphIndex, characterStart, characterEnd)) || [];
			const first = actions[0];
			if (!first?.barnOrientation || !first.barnMotion) return {};
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
			if (props.thumbnail && !props.playbackPreview || !actions?.length) return {};
			const visibilityActions = actions.filter((action) => [
				"appear",
				"fade",
				"wipe",
				"strips",
				"barn",
				"slide",
				"visibility",
				"zoom"
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
					if (action.visibilityKeyframes?.length) for (const frame of action.visibilityKeyframes) {
						if (frame.offset > progress) break;
						value = frame.value;
					}
					else if (progress >= 1) value = action.visibilityTo || value;
					visibility = value;
				} else if (action.effect === "wipe" || action.effect === "strips" || action.effect === "barn") {
					visibility = action.direction === "in" ? progress > 0 ? "visible" : "hidden" : progress >= 1 ? "hidden" : "visible";
					opacity = 1;
				} else if (action.effect === "slide" || action.effect === "zoom") {
					visibility = action.direction === "in" || progress < 1 ? "visible" : "hidden";
					opacity = 1;
				} else {
					opacity = action.direction === "in" ? progress : 1 - progress;
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
			if (!actions?.length) return {};
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
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
			const actions = animationsByTarget.value.get(element.id)?.filter((action) => {
				const range = action.characterRange;
				const overlaps = range && start < range.end && end > range.start;
				if (action.iteration) return Boolean(overlaps || !range) && (iterationIndexFor(action, start, end) ?? -1) >= 0;
				return Boolean(overlaps);
			});
			const animatedStyle = animationCss(actions?.filter((action) => action.effect !== "barn"), start, end);
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
			return triggerTargetIds.value.has(elementId);
		}
		function triggerSequenceIsActive(index, sequence, now) {
			return sequence.steps.some((actions, stepIndex) => {
				const key = triggerKey(index, stepIndex);
				const startedAt = triggerStartedAt.get(key);
				const endedAt = triggerStepEndedAt.get(key);
				if (startedAt === void 0 || !Number.isFinite(startedAt)) return false;
				if (endedAt !== void 0 && now >= endedAt) return false;
				return actions.some((action) => {
					if (!playbackActionEnabled(action)) return false;
					const actionStart = startedAt + action.delayMs;
					if (now < actionStart) return true;
					if (action.effect === "media" && action.mediaActionKey) {
						if (completedMediaActions.has(action.mediaActionKey)) return false;
						if (action.mediaWaitForEnd) {
							if (blockedMediaPlaybackIds.value.includes(action.targetId)) return false;
							const media = Array.from(stage.value?.querySelectorAll("[data-media-id]") || []).find((item) => item.dataset.mediaId === action.targetId);
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
				if (!sequence || !triggerSequenceHasEnabledAction(sequence)) continue;
				const count = sequence.steps.length;
				if (!count) continue;
				let step = triggerSteps.get(index) || 0;
				if (step >= count) {
					if (sequence.restart === "never") continue;
					if (sequence.restart === "whenNotActive" && triggerSequenceIsActive(index, sequence, performance.now())) continue;
					for (let previous = 0; previous < count; previous++) {
						const key = triggerKey(index, previous);
						triggerStartedAt.delete(key);
						triggerStepEndedAt.delete(key);
						triggerPlaybackByKey.delete(key);
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
			if (props.thumbnail || props.triggerPlayback !== void 0) return Infinity;
			let next = Infinity;
			props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
				if (!triggerSequenceHasEnabledAction(sequence)) return;
				sequence.conditions.forEach((condition, conditionIndex) => {
					const key = triggerConditionKey(sequenceIndex, conditionIndex);
					if (condition.type === "timeNode" && !firedTimingNodeTriggers.has(key)) {
						const mediaActionTrigger = condition.sourceMediaActionKey !== void 0;
						const sourceStartedAt = mediaActionTrigger ? void 0 : condition.sourceStep === void 0 ? automaticStartedAt : stepStartedAt.get(condition.sourceStep);
						const mediaEventAt = mediaActionTrigger ? condition.event === "begin" ? scheduledMediaActions.get(condition.sourceMediaActionKey) : condition.event === "onBegin" ? startedMediaActions.get(condition.sourceMediaActionKey) : condition.event === "end" ? completedMediaActions.get(condition.sourceMediaActionKey) : endedMediaTargets.get(condition.sourceMediaActionKey) : void 0;
						const sourceEndedAt = !mediaActionTrigger && condition.event === "onEnd" && condition.sourceStep !== void 0 ? mainStepEndedAt.get(condition.sourceStep) : void 0;
						const sourceTime = mediaActionTrigger ? mediaEventAt : sourceStartedAt !== void 0 && sourceStartedAt !== Number.NEGATIVE_INFINITY ? sourceEndedAt ?? sourceStartedAt + condition.sourceDelayMs : void 0;
						if (sourceTime !== void 0) {
							const startedAt = sourceTime + condition.delayMs;
							if (Number.isFinite(startedAt)) {
								if (startedAt > now) next = Math.min(next, startedAt);
								else {
									firedTimingNodeTriggers.add(key);
									triggerConditionDueAt.set(key, startedAt);
									pendingTriggerConditions.add(key);
								}
							}
						}
					}
					const dueAt = triggerConditionDueAt.get(key);
					if (pendingTriggerConditions.has(key) && dueAt !== void 0 && dueAt > now) next = Math.min(next, dueAt);
				});
				const pending = sequence.conditions.flatMap((_, conditionIndex) => {
					const key = triggerConditionKey(sequenceIndex, conditionIndex);
					const dueAt = triggerConditionDueAt.get(key);
					return pendingTriggerConditions.has(key) && dueAt !== void 0 && dueAt <= now ? [{
						conditionIndex,
						dueAt
					}] : [];
				});
				if (!pending.length) return;
				const progress = triggerSteps.get(sequenceIndex) || 0;
				const advancing = progress > 0 && progress < sequence.steps.length;
				const conditionTimes = sequence.conditions.map((_, conditionIndex) => triggerConditionDueAt.get(triggerConditionKey(sequenceIndex, conditionIndex)));
				const allConditionsDue = conditionTimes.every((dueAt) => dueAt !== void 0 && dueAt <= now);
				if (!advancing && !allConditionsDue) return;
				const startedAt = advancing ? Math.min(...pending.map((condition) => condition.dueAt)) : Math.max(...conditionTimes);
				sequence.conditions.forEach((condition, conditionIndex) => {
					const key = triggerConditionKey(sequenceIndex, conditionIndex);
					pendingTriggerConditions.delete(key);
					if (condition.type !== "timeNode") triggerConditionDueAt.delete(key);
				});
				startTriggeredAnimations([sequenceIndex], startedAt, false);
			});
			return next;
		}
		function satisfyTriggeredAnimationConditions(matches) {
			if (props.thumbnail || props.triggerPlayback !== void 0) return;
			const now = performance.now();
			let matched = false;
			props.slide.triggeredAnimations?.forEach((sequence, sequenceIndex) => {
				if (!triggerSequenceHasEnabledAction(sequence)) return;
				sequence.conditions.forEach((condition, conditionIndex) => {
					if (!matches(condition, sequenceIndex, conditionIndex)) return;
					matched = true;
					const key = triggerConditionKey(sequenceIndex, conditionIndex);
					if (triggerConditionDueAt.has(key)) return;
					triggerConditionDueAt.set(key, now + condition.delayMs);
					pendingTriggerConditions.add(key);
				});
			});
			if (matched) {
				animationClock.value = now;
				scheduleTimelineUpdate();
			}
		}
		function runSlideTriggeredAnimations(triggerEvent) {
			if (props.thumbnail || props.triggerPlayback !== void 0) return;
			satisfyTriggeredAnimationConditions((condition) => condition.type === "slide" && condition.event === triggerEvent);
		}
		function canActivateHyperlink(hyperlink) {
			return Boolean(hyperlink) && !props.thumbnail && !props.playbackPreview && props.triggerPlayback === void 0;
		}
		function hyperlinkHref(hyperlink) {
			return hyperlink?.kind === "external" ? hyperlink.url : "#slide-link";
		}
		function usesRightToLeftColumns(element) {
			return element.textColumnsRightToLeft === true && (element.textColumnCount ?? 1) > 1 && (!element.textOrientation || element.textOrientation === "horz");
		}
		function textFrameStyle(element) {
			const verticalText = element.textOrientation === "vert" || element.textOrientation === "vert270";
			const eastAsianVerticalText = element.textOrientation === "eaVert";
			const wordArtVerticalText = element.textOrientation === "wordArtVert";
			const wordArtVerticalRtlText = element.textOrientation === "wordArtVertRtl";
			const mongolianVerticalText = element.textOrientation === "mongolianVert";
			const columnCount = element.textColumnCount ?? 1;
			const verticalColumnWritingMode = columnCount > 1 && element.textOrientation && element.textOrientation !== "horz" && !wordArtVerticalRtlText ? element.textColumnsRightToLeft ? "vertical-rl" : "vertical-lr" : void 0;
			const rect = element.customTextRect;
			const fittedSize = rect ? autoFitDimensions.get(`${props.slide.id}:${element.id}`) : void 0;
			const ellipsisLineCount = supportsTextVerticalEllipsis(element) ? textEllipsisLineCounts.get(`${props.slide.id}:${element.id}`) : void 0;
			const margins = rect ? { ...element.margins } : element.margins;
			if (rect) {
				const availableWidth = Math.max(0, (rect.right - rect.left) * (fittedSize?.width || element.width));
				const horizontalMargins = margins.left + margins.right;
				if (horizontalMargins > availableWidth) {
					const ratio = availableWidth / horizontalMargins;
					margins.left *= ratio;
					margins.right *= ratio;
				}
				const availableHeight = Math.max(0, (rect.bottom - rect.top) * (fittedSize?.height || element.height));
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
				columnGap: columnCount > 1 ? `${(element.textColumnSpacing ?? 0) * scale.value}px` : void 0,
				columnFill: columnCount > 1 ? "auto" : void 0,
				padding: `${margins.top * scale.value}px ${margins.right * scale.value}px ${margins.bottom * scale.value}px ${margins.left * scale.value}px`,
				justifyContent: element.verticalAlign === "middle" ? "center" : element.verticalAlign === "bottom" ? "flex-end" : "flex-start",
				overflowX: element.textHorizontalOverflow === void 0 ? void 0 : element.textHorizontalOverflow === "clip" ? "clip" : "visible",
				overflowY: ellipsisLineCount ? "hidden" : element.textVerticalOverflow === void 0 ? void 0 : element.textVerticalOverflow === "overflow" ? "visible" : "clip",
				whiteSpace: element.textWrap === "none" ? "nowrap" : void 0,
				writingMode: verticalColumnWritingMode ?? (element.textOrientation === "vert270" || mongolianVerticalText ? "vertical-lr" : verticalText || eastAsianVerticalText || wordArtVerticalText || wordArtVerticalRtlText ? "vertical-rl" : void 0),
				textOrientation: verticalText ? "sideways" : eastAsianVerticalText || wordArtVerticalText || wordArtVerticalRtlText || mongolianVerticalText ? "upright" : void 0,
				direction: element.textOrientation === "vert270" || usesRightToLeftColumns(element) ? "rtl" : void 0
			};
		}
		function textBoundsStyle(element) {
			if (!element.textAnchorCenter) return { display: "contents" };
			const centeredOffset = textAnchorCenterOffsets.get(`${props.slide.id}:${element.id}`);
			return {
				display: (element.textColumnCount ?? 1) > 1 ? "block" : "flex",
				flexDirection: (element.textColumnCount ?? 1) > 1 ? void 0 : "column",
				width: "fit-content",
				maxWidth: "100%",
				marginLeft: "auto",
				marginRight: "auto",
				transform: centeredOffset ? `translateX(${centeredOffset}px)` : void 0
			};
		}
		function effectiveTextFontScale(element) {
			return normalAutoFitScales.get(`${props.slide.id}:${element.id}`) ?? element.textFontScale ?? 1;
		}
		function tableCellAutoFitKey(elementId, rowIndex, cellIndex) {
			return `${props.slide.id}:${elementId}:${rowIndex}:${cellIndex}`;
		}
		function effectiveTableCellFontScale(element, rowIndex, cellIndex, cell) {
			return normalAutoFitScales.get(tableCellAutoFitKey(element.id, rowIndex, cellIndex)) ?? cell.textFontScale ?? 1;
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
		function textInnerShadows(run) {
			return (run.shadows?.length ? run.shadows : run.shadow ? [run.shadow] : []).filter((shadow) => shadow.inset);
		}
		function textInnerShadowFilterId(element, paragraphIndex, characterStart, characterEnd, filterKey) {
			return `${customGradientScope}-text-inner-shadow-${element.id}-${filterKey}-${paragraphIndex}-${characterStart}-${characterEnd}`.replace(/[^a-zA-Z0-9_-]/g, "_");
		}
		function textInnerShadowFilterReference(element, run, paragraphIndex, characterStart, characterEnd, filterKey) {
			return textInnerShadows(run).length ? `url(#${textInnerShadowFilterId(element, paragraphIndex, characterStart, characterEnd, filterKey)})` : void 0;
		}
		function textInnerShadowOffset(shadow) {
			const angle = (shadow.angle || 0) * Math.PI / 180;
			return {
				dx: Math.cos(angle) * shadow.distance * scale.value,
				dy: Math.sin(angle) * shadow.distance * scale.value
			};
		}
		function textInnerShadowColor(element, shadow, paragraphIndex, characterStart, characterEnd) {
			const color = animationColor(element, "shadow.color", paragraphIndex, characterStart, characterEnd) || shadow.color;
			return animatedRunColor(color, animationOpacity(element, "shadow.opacity", paragraphIndex, characterStart, characterEnd)) || color;
		}
		function runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd) {
			const shadows = (run.shadows?.length ? run.shadows : run.shadow ? [run.shadow] : []).filter((shadow) => !shadow.inset);
			const glows = run.glows?.length ? run.glows : run.glow ? [run.glow] : [];
			const effects = shadows.map((shadow) => {
				const shadowColor = animationColor(element, "shadow.color", paragraphIndex, characterStart, characterEnd) || shadow.color;
				return shadowCss(shadow, animatedRunColor(shadowColor, animationOpacity(element, "shadow.opacity", paragraphIndex, characterStart, characterEnd)) || shadowColor);
			});
			effects.push(...glows.map((glow) => `0 0 ${glow.radius * scale.value}px ${glow.color}`));
			return effects.length ? effects.join(", ") : void 0;
		}
		function cssBackgroundLayerCount(background) {
			let depth = 0;
			let layers = 1;
			for (const character of background) if (character === "(") depth++;
			else if (character === ")") depth--;
			else if (character === "," && depth === 0) layers++;
			return layers;
		}
		function runStyle(element, run, script, fontScale, paragraphIndex, characterStart, characterEnd, separateDecorations = false, filterKey = `${paragraphIndex}-${characterStart}-${characterEnd}`) {
			const requestedFont = script === "eastAsia" ? run.fontFamilyEastAsia : script === "complexScript" ? run.fontFamilyComplexScript : run.fontFamily;
			const fillOpacity = animationOpacity(element, "fill.opacity", paragraphIndex, characterStart, characterEnd);
			const strokeOpacity = animationOpacity(element, "stroke.opacity", paragraphIndex, characterStart, characterEnd);
			const textFill = run.fillGradient ? applyColorOpacity(run.fillGradient, fillOpacity) : run.fillPattern ? applyColorOpacity(run.fillPattern, fillOpacity) : void 0;
			const highlightColor = run.highlightColor ? applyColorOpacity(run.highlightColor, fillOpacity) : void 0;
			const highlightBackground = textFill && highlightColor ? `linear-gradient(${highlightColor}, ${highlightColor}) border-box, ${textFill}` : void 0;
			const textFillClip = textFill && highlightColor ? `border-box, ${Array(cssBackgroundLayerCount(textFill)).fill("text").join(", ")}` : void 0;
			const textDecorationLine = separateDecorations ? run.lineThrough ? "line-through" : void 0 : [run.underline ? "underline" : void 0, run.lineThrough ? "line-through" : void 0].filter(Boolean).join(" ");
			const textDecorationStyle = separateDecorations ? run.lineThroughStyle : run.underline && run.lineThrough ? run.underlineStyle ?? (run.lineThroughStyle === "double" ? "solid" : run.lineThroughStyle) : run.underlineStyle ?? run.lineThroughStyle;
			return {
				fontFamily: cssFontFamily(requestedFont || run.fontFamily),
				fontSize: `${(run.fontSizePt || 18) * fontScale * EMU_PER_POINT * scale.value}px`,
				letterSpacing: run.letterSpacingPt === void 0 ? void 0 : `${run.letterSpacingPt * fontScale * EMU_PER_POINT * scale.value}px`,
				verticalAlign: run.baselineOffsetEm === void 0 ? void 0 : `${run.baselineOffsetEm}em`,
				color: animatedRunColor(run.color, fillOpacity),
				"--pptx-text-fill": textFill,
				"--pptx-text-fill-color": animatedRunColor(run.color, fillOpacity),
				"--pptx-text-highlight": highlightColor,
				"--pptx-text-highlight-background": highlightBackground,
				"--pptx-text-fill-clip": textFillClip,
				WebkitTextStroke: run.outlineColor !== void 0 && run.outlineWidth !== void 0 ? `${run.outlineWidth * scale.value}px ${animatedRunColor(run.outlineColor, strokeOpacity)}` : void 0,
				fontWeight: run.bold ? 700 : void 0,
				fontStyle: run.italic ? "italic" : void 0,
				...separateDecorations ? { "--pptx-underline-line": run.underline ? "underline" : "none" } : {},
				textDecorationLine: textDecorationLine || void 0,
				textDecorationStyle,
				textDecorationColor: !separateDecorations && run.underline ? run.underlineColor : void 0,
				textDecorationThickness: !separateDecorations && run.underline && run.underlineThicknessEmu ? `${run.underlineThicknessEmu * scale.value}px` : void 0,
				textShadow: runTextShadowStyle(element, run, paragraphIndex, characterStart, characterEnd),
				filter: textInnerShadowFilterReference(element, run, paragraphIndex, characterStart, characterEnd, filterKey)
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
			let offset = 0;
			for (let index = 0; index < paragraphIndex; index++) {
				const paragraph = element.paragraphs[index];
				if (paragraph) offset += paragraph.runs.reduce((sum, run) => sum + run.text.length, 0) + 1;
			}
			const paragraph = element.paragraphs[paragraphIndex];
			if (paragraph) for (let index = 0; index < runIndex; index++) offset += paragraph.runs[index]?.text.length || 0;
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
			const runStart = runCharacterOffset(element, paragraphIndex, runIndex);
			const characterRanges = (animationsByTarget.value.get(element.id) || []).flatMap((action) => [...action.characterRange ? [action.characterRange] : [], ...action.iteration?.ranges || []]);
			if (!characterRanges.length) {
				const segments = [];
				let segmentStart = 0;
				let segmentScript;
				let previousScript = "latin";
				let offset = 0;
				for (const character of run.text) {
					const script = characterScript(character, previousScript);
					if (segmentScript && script !== segmentScript) {
						segments.push({
							text: run.text.slice(segmentStart, offset),
							script: segmentScript,
							start: runStart + segmentStart,
							end: runStart + offset
						});
						segmentStart = offset;
					}
					segmentScript = script;
					previousScript = script;
					offset += character.length;
				}
				if (segmentScript && offset > segmentStart) segments.push({
					text: run.text.slice(segmentStart),
					script: segmentScript,
					start: runStart + segmentStart,
					end: runStart + offset
				});
				return segments.map((segment) => ({
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
				if (characters.length && characters.at(-1)?.script !== script) cuts.add(offset);
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
			const size = paragraph.bullet?.size;
			if (!size) return {};
			return { fontSize: size.unit === "ratio" ? `${size.value * 100}%` : `${size.value * EMU_PER_POINT * scale.value}px` };
		}
		function cssSpacing(spacing) {
			if (!spacing) return void 0;
			return spacing.unit === "emu" ? `${spacing.value * scale.value}px` : `${spacing.value}em`;
		}
		function paragraphStyle(paragraph, fontScale = 1, lineSpaceReduction = 0, fallbackDirection) {
			const lineSpacing = paragraph.lineSpacing;
			return {
				fontSize: `${(paragraph.runs.reduce((size, run) => Math.max(size, run.fontSizePt || 18), 0) || 18) * fontScale * EMU_PER_POINT * scale.value}px`,
				textAlign: paragraph.align,
				direction: paragraph.direction ?? fallbackDirection,
				tabSize: paragraph.tabStopInterval === void 0 ? void 0 : `${paragraph.tabStopInterval * scale.value}px`,
				marginLeft: paragraph.marginLeft === void 0 ? void 0 : `${paragraph.marginLeft * scale.value}px`,
				marginRight: paragraph.marginRight === void 0 ? void 0 : `${paragraph.marginRight * scale.value}px`,
				textIndent: paragraph.indent === void 0 ? void 0 : `${paragraph.indent * scale.value}px`,
				marginTop: cssSpacing(paragraph.spaceBefore),
				marginBottom: cssSpacing(paragraph.spaceAfter),
				lineHeight: lineSpacing?.unit === "ratio" ? lineSpacing.value * (1 - lineSpaceReduction) : cssSpacing(lineSpacing)
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
			const columnCount = cell.textColumnCount ?? 1;
			return {
				...columnCount > 1 || cell.textAutoFitDynamic ? {
					height: rowHeight > 0 ? `${Math.max(0, rowHeight - cell.margins.top - cell.margins.bottom) * scale.value}px` : "100%",
					boxSizing: "border-box",
					minHeight: 0,
					overflow: "hidden",
					...columnCount > 1 ? {
						columnCount,
						columnGap: `${(cell.textColumnSpacing ?? 0) * scale.value}px`,
						columnFill: "auto",
						direction: cell.textColumnsRightToLeft ? "rtl" : void 0
					} : {}
				} : {},
				...cell.textAnchorCenter ? {
					width: "fit-content",
					maxWidth: "100%",
					marginLeft: "auto",
					marginRight: "auto"
				} : {}
			};
		}
		return (_ctx, _push, _parent, _attrs) => {
			_push(`<div${ssrRenderAttrs(mergeProps({
				ref_key: "stage",
				ref: stage,
				class: ["slide-frame", { "is-thumbnail": __props.thumbnail }],
				style: frameStyle.value,
				role: hasInteractiveMedia.value ? "group" : "img",
				"aria-label": `Slide ${__props.number}`
			}, _attrs))} data-v-954c30ab><!--[-->`);
			ssrRenderList(__props.slide.elements, (element) => {
				_push(`<div${ssrRenderAttr("data-element-id", element.id)}${ssrRenderAttr("data-element-name", element.name)}${ssrRenderAttr("data-element-kind", element.kind)} class="${ssrRenderClass([{
					"slide-shape": element.kind === "shape",
					"shape-line": element.kind === "line",
					"slide-placeholder": element.kind === "placeholder",
					"is-animation-trigger": hasTrigger(element.id),
					"is-hyperlink": canActivateHyperlink(element.hyperlink)
				}, "slide-element"])}"${ssrRenderAttr("role", canActivateHyperlink(element.hyperlink) ? "link" : hasTrigger(element.id) ? "button" : void 0)}${ssrRenderAttr("tabindex", canActivateHyperlink(element.hyperlink) || hasTrigger(element.id) ? 0 : void 0)}${ssrRenderAttr("aria-label", element.hyperlink?.tooltip || (hasTrigger(element.id) ? `Play animation from ${element.name}` : void 0))}${ssrRenderAttr("data-pptx-hyperlink", canActivateHyperlink(element.hyperlink) ? JSON.stringify(element.hyperlink) : void 0)} style="${ssrRenderStyle([
					elementStyle(element),
					geometryStyle(element),
					animationStyle(element)
				])}"${ssrRenderAttr("title", element.hyperlink?.tooltip || element.name)} data-v-954c30ab>`);
				if (usesSvgOutline(element)) {
					_push(`<svg class="preset-outline" viewBox="0 0 1000 1000" preserveAspectRatio="none" style="${ssrRenderStyle(usesSvgPolygon(element) ? customGeometryStyle(element) : void 0)}" aria-hidden="true" data-v-954c30ab><defs data-v-954c30ab>`);
					if (usesSvgPolygon(element) && customGeometryInnerShadows(element).length && animationOnOff(element, "shadow.on")) {
						_push(`<filter${ssrRenderAttrs(mergeProps({ id: customGeometryInnerShadowFilterId(element) }, { ref_for: true }, customGeometryInnerShadowFilterRegion(element), {
							filterUnits: "objectBoundingBox",
							primitiveUnits: "objectBoundingBox",
							"color-interpolation-filters": "sRGB"
						}))} data-v-954c30ab><!--[-->`);
						ssrRenderList(customGeometryInnerShadows(element), (shadow, shadowIndex) => {
							_push(`<!--[--><feGaussianBlur in="SourceAlpha"${ssrRenderAttr("stdDeviation", customGeometryInnerShadowDeviation(element, shadow))}${ssrRenderAttr("result", "custom-inner-blur-" + shadowIndex)} data-v-954c30ab></feGaussianBlur><feOffset${ssrRenderAttr("in", "custom-inner-blur-" + shadowIndex)}${ssrRenderAttr("dx", customGeometryInnerShadowOffset(element, shadow).dx)}${ssrRenderAttr("dy", customGeometryInnerShadowOffset(element, shadow).dy)}${ssrRenderAttr("result", "custom-inner-offset-" + shadowIndex)} data-v-954c30ab></feOffset><feComposite in="SourceAlpha"${ssrRenderAttr("in2", "custom-inner-offset-" + shadowIndex)} operator="out"${ssrRenderAttr("result", "custom-inner-mask-" + shadowIndex)} data-v-954c30ab></feComposite><feFlood${ssrRenderAttr("flood-color", customGeometryInnerShadowColor(element, shadow))}${ssrRenderAttr("result", "custom-inner-color-" + shadowIndex)} data-v-954c30ab></feFlood><feComposite${ssrRenderAttr("in", "custom-inner-color-" + shadowIndex)}${ssrRenderAttr("in2", "custom-inner-mask-" + shadowIndex)} operator="in"${ssrRenderAttr("result", "custom-inner-shadow-" + shadowIndex)} data-v-954c30ab></feComposite><!--]-->`);
						});
						_push(`<!--]--><feMerge data-v-954c30ab><!--[-->`);
						ssrRenderList(customGeometryInnerShadows(element), (_, shadowIndex) => {
							_push(`<feMergeNode${ssrRenderAttr("in", "custom-inner-shadow-" + shadowIndex)} data-v-954c30ab></feMergeNode>`);
						});
						_push(`<!--]--><feMergeNode in="SourceGraphic" data-v-954c30ab></feMergeNode></feMerge></filter>`);
					} else _push(`<!---->`);
					_push(`<!--[-->`);
					ssrRenderList(customGeometryGradients(element).filter((item) => usesSvgPolygon(element) || item.id === customGradientId(element, "stroke")), (gradient) => {
						_push(`<!--[-->`);
						if (gradient.kind === "linear") {
							_push(`<linearGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("x1", gradient.x1)}${ssrRenderAttr("y1", gradient.y1)}${ssrRenderAttr("x2", gradient.x2)}${ssrRenderAttr("y2", gradient.y2)} data-v-954c30ab><!--[-->`);
							ssrRenderList(gradient.stops, (stop, index) => {
								_push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-954c30ab></stop>`);
							});
							_push(`<!--]--></linearGradient>`);
						} else {
							_push(`<radialGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("cx", gradient.cx)}${ssrRenderAttr("cy", gradient.cy)}${ssrRenderAttr("r", gradient.r)}${ssrRenderAttr("fx", gradient.fx)}${ssrRenderAttr("fy", gradient.fy)}${ssrRenderAttr("fr", gradient.fr)}${ssrRenderAttr("gradientTransform", gradient.gradientTransform)}${ssrRenderAttr("color-interpolation", gradient.colorInterpolation)} data-v-954c30ab><!--[-->`);
							ssrRenderList(gradient.stops, (stop, index) => {
								_push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-954c30ab></stop>`);
							});
							_push(`<!--]--></radialGradient>`);
						}
						_push(`<!--]-->`);
					});
					_push(`<!--]--><!--[-->`);
					ssrRenderList(customGeometryPatternModes(element), (mode) => {
						_push(`<pattern${ssrRenderAttr("id", customGeometryPatternId(element, mode))} patternUnits="userSpaceOnUse" width="1000" height="1000" data-v-954c30ab><foreignObject width="1000" height="1000" data-v-954c30ab><div xmlns="http://www.w3.org/1999/xhtml" style="${ssrRenderStyle([{
							"width": "100%",
							"height": "100%"
						}, { background: applyColorOpacity(element.fill, 1, customPathBrightness(mode)) }])}" data-v-954c30ab></div></foreignObject></pattern>`);
					});
					_push(`<!--]-->`);
					if (element.strokePattern) _push(`<pattern${ssrRenderAttr("id", customGeometryStrokePatternId(element))} patternUnits="userSpaceOnUse" width="1000" height="1000" data-v-954c30ab><foreignObject width="1000" height="1000" data-v-954c30ab><div xmlns="http://www.w3.org/1999/xhtml" style="${ssrRenderStyle([{
						"width": "100%",
						"height": "100%"
					}, { background: element.strokePattern }])}" data-v-954c30ab></div></foreignObject></pattern>`);
					else _push(`<!---->`);
					_push(`</defs><path${ssrRenderAttr("d", presetOutlinePath(element))}${ssrRenderAttr("fill", usesSvgPolygon(element) ? customGeometryFill(element, {
						fill: true}) : "none")}${ssrRenderAttr("fill-opacity", usesSvgPolygon(element) ? customGeometryFillOpacity(element, {
						fill: true}) : void 0)}${ssrRenderAttr("stroke", presetOutlineDetailPaths(element).length ? "none" : customGeometryStroke(element, {
						stroke: true
					}))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}" vector-effect="non-scaling-stroke" data-v-954c30ab></path><!--[-->`);
					ssrRenderList(presetOutlineDetailPaths(element), (detailPath, index) => {
						_push(`<path${ssrRenderAttr("d", detailPath)} fill="none"${ssrRenderAttr("stroke", customGeometryStroke(element, {
							stroke: true
						}))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}" vector-effect="non-scaling-stroke" data-v-954c30ab></path>`);
					});
					_push(`<!--]--></svg>`);
				} else _push(`<!---->`);
				if (element.kind === "shape" && element.customPaths?.length) {
					_push(`<svg class="custom-geometry" viewBox="0 0 1000 1000" preserveAspectRatio="none" style="${ssrRenderStyle(customGeometryStyle(element))}" aria-hidden="true" data-v-954c30ab><defs data-v-954c30ab>`);
					if (customGeometryInnerShadows(element).length && animationOnOff(element, "shadow.on")) {
						_push(`<filter${ssrRenderAttrs(mergeProps({ id: customGeometryInnerShadowFilterId(element) }, { ref_for: true }, customGeometryInnerShadowFilterRegion(element), {
							filterUnits: "objectBoundingBox",
							primitiveUnits: "objectBoundingBox",
							"color-interpolation-filters": "sRGB"
						}))} data-v-954c30ab><!--[-->`);
						ssrRenderList(customGeometryInnerShadows(element), (shadow, shadowIndex) => {
							_push(`<!--[--><feGaussianBlur in="SourceAlpha"${ssrRenderAttr("stdDeviation", customGeometryInnerShadowDeviation(element, shadow))}${ssrRenderAttr("result", "custom-inner-blur-" + shadowIndex)} data-v-954c30ab></feGaussianBlur><feOffset${ssrRenderAttr("in", "custom-inner-blur-" + shadowIndex)}${ssrRenderAttr("dx", customGeometryInnerShadowOffset(element, shadow).dx)}${ssrRenderAttr("dy", customGeometryInnerShadowOffset(element, shadow).dy)}${ssrRenderAttr("result", "custom-inner-offset-" + shadowIndex)} data-v-954c30ab></feOffset><feComposite in="SourceAlpha"${ssrRenderAttr("in2", "custom-inner-offset-" + shadowIndex)} operator="out"${ssrRenderAttr("result", "custom-inner-mask-" + shadowIndex)} data-v-954c30ab></feComposite><feFlood${ssrRenderAttr("flood-color", customGeometryInnerShadowColor(element, shadow))}${ssrRenderAttr("result", "custom-inner-color-" + shadowIndex)} data-v-954c30ab></feFlood><feComposite${ssrRenderAttr("in", "custom-inner-color-" + shadowIndex)}${ssrRenderAttr("in2", "custom-inner-mask-" + shadowIndex)} operator="in"${ssrRenderAttr("result", "custom-inner-shadow-" + shadowIndex)} data-v-954c30ab></feComposite><!--]-->`);
						});
						_push(`<!--]--><feMerge data-v-954c30ab><!--[-->`);
						ssrRenderList(customGeometryInnerShadows(element), (_, shadowIndex) => {
							_push(`<feMergeNode${ssrRenderAttr("in", "custom-inner-shadow-" + shadowIndex)} data-v-954c30ab></feMergeNode>`);
						});
						_push(`<!--]--><feMergeNode in="SourceGraphic" data-v-954c30ab></feMergeNode></feMerge></filter>`);
					} else _push(`<!---->`);
					_push(`<!--[-->`);
					ssrRenderList(customGeometryGradients(element), (gradient) => {
						_push(`<!--[-->`);
						if (gradient.kind === "linear") {
							_push(`<linearGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("x1", gradient.x1)}${ssrRenderAttr("y1", gradient.y1)}${ssrRenderAttr("x2", gradient.x2)}${ssrRenderAttr("y2", gradient.y2)} data-v-954c30ab><!--[-->`);
							ssrRenderList(gradient.stops, (stop, stopIndex) => {
								_push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-954c30ab></stop>`);
							});
							_push(`<!--]--></linearGradient>`);
						} else {
							_push(`<radialGradient${ssrRenderAttr("id", gradient.id)} gradientUnits="userSpaceOnUse"${ssrRenderAttr("cx", gradient.cx)}${ssrRenderAttr("cy", gradient.cy)}${ssrRenderAttr("r", gradient.r)}${ssrRenderAttr("fx", gradient.fx)}${ssrRenderAttr("fy", gradient.fy)}${ssrRenderAttr("fr", gradient.fr)}${ssrRenderAttr("gradientTransform", gradient.gradientTransform)}${ssrRenderAttr("color-interpolation", gradient.colorInterpolation)} data-v-954c30ab><!--[-->`);
							ssrRenderList(gradient.stops, (stop, stopIndex) => {
								_push(`<stop${ssrRenderAttr("offset", stop.offset)}${ssrRenderAttr("stop-color", stop.color)}${ssrRenderAttr("stop-opacity", stop.opacity)} data-v-954c30ab></stop>`);
							});
							_push(`<!--]--></radialGradient>`);
						}
						_push(`<!--]-->`);
					});
					_push(`<!--]--><!--[-->`);
					ssrRenderList(customGeometryPatternModes(element), (mode) => {
						_push(`<pattern${ssrRenderAttr("id", customGeometryPatternId(element, mode))} patternUnits="userSpaceOnUse" width="1000" height="1000" data-v-954c30ab><foreignObject width="1000" height="1000" data-v-954c30ab><div xmlns="http://www.w3.org/1999/xhtml" style="${ssrRenderStyle([{
							"width": "100%",
							"height": "100%"
						}, { background: applyColorOpacity(element.fill, 1, customPathBrightness(mode)) }])}" data-v-954c30ab></div></foreignObject></pattern>`);
					});
					_push(`<!--]-->`);
					if (element.strokePattern) _push(`<pattern${ssrRenderAttr("id", customGeometryStrokePatternId(element))} patternUnits="userSpaceOnUse" width="1000" height="1000" data-v-954c30ab><foreignObject width="1000" height="1000" data-v-954c30ab><div xmlns="http://www.w3.org/1999/xhtml" style="${ssrRenderStyle([{
						"width": "100%",
						"height": "100%"
					}, { background: element.strokePattern }])}" data-v-954c30ab></div></foreignObject></pattern>`);
					else _push(`<!---->`);
					_push(`<!--[-->`);
					ssrRenderList(element.customPaths, (path, pathIndex) => {
						_push(`<!--[-->`);
						if (customGeometryFillFocus(element, path)) _push(`<clipPath${ssrRenderAttr("id", customGradientFocusClipId(element, pathIndex))} clipPathUnits="userSpaceOnUse" data-v-954c30ab><path${ssrRenderAttr("d", path.d)} data-v-954c30ab></path></clipPath>`);
						else _push(`<!---->`);
						_push(`<!--]-->`);
					});
					_push(`<!--]--></defs><!--[-->`);
					ssrRenderList(element.customPaths, (path, pathIndex) => {
						_push(`<!--[--><path${ssrRenderAttr("d", path.d)}${ssrRenderAttr("fill", customGeometryFill(element, path))}${ssrRenderAttr("fill-opacity", customGeometryFillOpacity(element, path))}${ssrRenderAttr("stroke", customGeometryStroke(element, path))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}" vector-effect="non-scaling-stroke" data-v-954c30ab></path><!--[-->`);
						ssrRenderList([customGeometryFillFocus(element, path)], (focus) => {
							_push(`<!--[-->`);
							if (focus?.kind === "ellipse") _push(`<ellipse${ssrRenderAttr("cx", focus.cx)}${ssrRenderAttr("cy", focus.cy)}${ssrRenderAttr("rx", focus.rx)}${ssrRenderAttr("ry", focus.ry)}${ssrRenderAttr("fill", focus.color)}${ssrRenderAttr("clip-path", `url(#${customGradientFocusClipId(element, pathIndex)})`)} data-v-954c30ab></ellipse>`);
							else if (focus) _push(`<rect${ssrRenderAttr("x", focus.x)}${ssrRenderAttr("y", focus.y)}${ssrRenderAttr("width", focus.width)}${ssrRenderAttr("height", focus.height)}${ssrRenderAttr("fill", focus.color)}${ssrRenderAttr("clip-path", `url(#${customGradientFocusClipId(element, pathIndex)})`)} data-v-954c30ab></rect>`);
							else _push(`<!---->`);
							_push(`<!--]-->`);
						});
						_push(`<!--]--><!--]-->`);
					});
					_push(`<!--]--></svg>`);
				} else _push(`<!---->`);
				if (element.kind === "video" && element.mediaUrl && !__props.thumbnail) {
					_push(`<div class="image-viewport" data-v-954c30ab><video class="slide-video"${ssrRenderAttr("data-media-id", element.id)}${ssrRenderAttr("data-media-slide-count", mediaSlideCounts.value.get(element.id) || void 0)}${ssrRenderAttr("src", element.mediaUrl)}${ssrRenderAttr("poster", element.imageUrl)} controls playsinline preload="metadata"${ssrRenderAttr("aria-label", element.name)} style="${ssrRenderStyle(imageStyle(element))}" data-v-954c30ab></video>`);
					if (invalidMediaTrimIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-954c30ab>Trim range exceeds media duration; full media is used.</span>`);
					else _push(`<!---->`);
					if (invalidMediaFadeIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-954c30ab>Media fade duration exceeds the playable range; fades are ignored.</span>`);
					else _push(`<!---->`);
					if (blockedMediaPlaybackIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-954c30ab>The browser blocked automatic playback; use the media controls.</span>`);
					else _push(`<!---->`);
					_push(`</div>`);
				} else if (element.kind === "audio" && element.mediaUrl && !__props.thumbnail) {
					_push(`<div class="audio-viewport" data-v-954c30ab><audio class="slide-audio"${ssrRenderAttr("data-media-id", element.id)}${ssrRenderAttr("data-media-slide-count", mediaSlideCounts.value.get(element.id) || void 0)}${ssrRenderAttr("src", element.mediaUrl)} controls preload="metadata"${ssrRenderAttr("aria-label", element.name)} data-v-954c30ab></audio>`);
					if (invalidMediaTrimIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-954c30ab>Trim range exceeds media duration; full media is used.</span>`);
					else _push(`<!---->`);
					if (invalidMediaFadeIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-954c30ab>Media fade duration exceeds the playable range; fades are ignored.</span>`);
					else _push(`<!---->`);
					if (blockedMediaPlaybackIds.value.includes(element.id)) _push(`<span class="media-warning" role="status" data-v-954c30ab>The browser blocked automatic playback; use the media controls.</span>`);
					else _push(`<!---->`);
					_push(`</div>`);
				} else if (element.kind === "chart" && element.chart) {
					_push(`<svg class="slide-chart"${ssrRenderAttr("data-chart-type", element.chart.type)}${ssrRenderAttr("data-radar-style", element.chart.radarStyle)} viewBox="0 0 1000 600" preserveAspectRatio="none" role="img"${ssrRenderAttr("aria-label", element.chart.title || element.name)} data-v-954c30ab><title data-v-954c30ab>${ssrInterpolate(element.chart.title || element.name)}</title><defs data-v-954c30ab><clipPath${ssrRenderAttr("id", `chart-plot-${element.id}`)} data-v-954c30ab><polygon${ssrRenderAttr("points", `${chartPlot(element.id).left},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top} ${chartPlot(element.id).left + chartPlot(element.id).width},${chartPlot(element.id).top + chartPlot(element.id).height} ${chartPlot(element.id).left},${chartPlot(element.id).top + chartPlot(element.id).height}`)} data-v-954c30ab></polygon></clipPath>`);
					if (chartVisuals.value.get(element.id)?.legendClip) _push(`<clipPath${ssrRenderAttr("id", chartVisuals.value.get(element.id)?.legendClip?.id)} data-v-954c30ab><rect${ssrRenderAttr("x", chartVisuals.value.get(element.id)?.legendClip?.x)}${ssrRenderAttr("y", chartVisuals.value.get(element.id)?.legendClip?.y)}${ssrRenderAttr("width", chartVisuals.value.get(element.id)?.legendClip?.width)}${ssrRenderAttr("height", chartVisuals.value.get(element.id)?.legendClip?.height)} data-v-954c30ab></rect></clipPath>`);
					else _push(`<!---->`);
					_push(`<!--[-->`);
					ssrRenderList(chartPatternDefinitions.value.get(element.id), (pattern) => {
						_push(`<pattern${ssrRenderAttr("id", pattern.id)} patternUnits="userSpaceOnUse" width="1000" height="600" data-v-954c30ab><foreignObject width="1000" height="600" data-v-954c30ab><div xmlns="http://www.w3.org/1999/xhtml" style="${ssrRenderStyle([{
							"width": "100%",
							"height": "100%"
						}, { background: pattern.fill }])}" data-v-954c30ab></div></foreignObject></pattern>`);
					});
					_push(`<!--]--></defs>`);
					if (element.chart.title) _push(`<text x="500" y="42" text-anchor="middle" class="chart-title" data-v-954c30ab>${ssrInterpolate(element.chart.title)}</text>`);
					else _push(`<!---->`);
					if (element.chart.type === "pie" || element.chart.type === "doughnut") {
						_push(`<g data-v-954c30ab><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.slices, (slice, index) => {
							_push(`<path${ssrRenderAttr("d", slice.d)}${ssrRenderAttr("fill", chartPaint(element.id, slice.color))}${ssrRenderAttr("fill-rule", element.chart.type === "doughnut" ? "evenodd" : "nonzero")}${ssrRenderAttr("stroke", chartPaint(element.id, slice.stroke))}${ssrRenderAttr("stroke-width", slice.strokeWidth ?? 2)}${ssrRenderAttr("stroke-dasharray", slice.strokeDasharray)}${ssrRenderAttr("stroke-linecap", slice.strokeLinecap)} data-v-954c30ab></path>`);
						});
						_push(`<!--]--></g>`);
					} else {
						_push(`<g data-v-954c30ab>`);
						if (element.chart.type === "radar") {
							_push(`<!--[--><!--[-->`);
							ssrRenderList(chartVisuals.value.get(element.id)?.radarGrid, (grid, index) => {
								_push(`<path class="chart-major-gridline chart-radar-gridline"${ssrRenderAttr("d", grid.d)} fill="none" stroke="#d1d5db" stroke-width="1" data-v-954c30ab></path>`);
							});
							_push(`<!--]--><!--[-->`);
							ssrRenderList(chartVisuals.value.get(element.id)?.radarSpokes, (spoke, index) => {
								_push(`<line${ssrRenderAttrs(mergeProps({
									key: `radar-spoke-${index}`,
									class: "chart-major-gridline chart-radar-spoke"
								}, { ref_for: true }, spoke, {
									stroke: "#d1d5db",
									"stroke-width": "1"
								}))} data-v-954c30ab></line>`);
							});
							_push(`<!--]--><!--]-->`);
						} else {
							_push(`<!--[-->`);
							if (!element.chart.categoryAxisDeleted && element.chart.categoryAxisMajorGridlines) {
								_push(`<!--[-->`);
								ssrRenderList(chartVisuals.value.get(element.id)?.categoryGridlines, (grid, index) => {
									_push(`<line${ssrRenderAttrs(mergeProps({
										key: `category-grid-${index}`,
										class: "chart-major-gridline chart-category-major-gridline",
										"data-axis": "category"
									}, { ref_for: true }, grid, {
										stroke: "#d1d5db",
										"stroke-width": "1"
									}))} data-v-954c30ab></line>`);
								});
								_push(`<!--]-->`);
							} else _push(`<!---->`);
							if (!element.chart.valueAxis?.deleted && element.chart.valueAxis?.majorGridlines) {
								_push(`<!--[-->`);
								ssrRenderList(chartVisuals.value.get(element.id)?.ticks, (tick, index) => {
									_push(`<line class="chart-major-gridline"${ssrRenderAttr("x1", element.chart.direction === "horizontal" ? tick.x : chartPlot(element.id).left)}${ssrRenderAttr("x2", element.chart.direction === "horizontal" ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width)}${ssrRenderAttr("y1", element.chart.direction === "horizontal" ? chartPlot(element.id).top : tick.y)}${ssrRenderAttr("y2", element.chart.direction === "horizontal" ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y)} stroke="#d1d5db" stroke-width="1" data-v-954c30ab></line>`);
								});
								_push(`<!--]-->`);
							} else _push(`<!---->`);
							_push(`<!--[-->`);
							ssrRenderList(chartVisuals.value.get(element.id)?.secondaryValueAxes, (axis) => {
								_push(`<!--[-->`);
								if (!axis.deleted && axis.majorGridlines) {
									_push(`<!--[-->`);
									ssrRenderList(axis.ticks, (tick, index) => {
										_push(`<line class="chart-major-gridline chart-secondary-gridline"${ssrRenderAttr("x1", axis.horizontal ? tick.x : chartPlot(element.id).left)}${ssrRenderAttr("x2", axis.horizontal ? tick.x : chartPlot(element.id).left + chartPlot(element.id).width)}${ssrRenderAttr("y1", axis.horizontal ? chartPlot(element.id).top : tick.y)}${ssrRenderAttr("y2", axis.horizontal ? chartPlot(element.id).top + chartPlot(element.id).height : tick.y)} stroke="#d1d5db" stroke-width="1" data-v-954c30ab></line>`);
									});
									_push(`<!--]-->`);
								} else _push(`<!---->`);
								_push(`<!--]-->`);
							});
							_push(`<!--]-->`);
							if (["scatter", "bubble"].includes(element.chart.type) && !element.chart.xAxis?.deleted && element.chart.xAxis?.majorGridlines) {
								_push(`<!--[-->`);
								ssrRenderList(chartVisuals.value.get(element.id)?.xTicks, (tick, index) => {
									_push(`<line class="chart-major-gridline chart-x-gridline"${ssrRenderAttr("x1", tick.x)}${ssrRenderAttr("x2", tick.x)}${ssrRenderAttr("y1", chartPlot(element.id).top)}${ssrRenderAttr("y2", chartPlot(element.id).top + chartPlot(element.id).height)} stroke="#d1d5db" stroke-width="1" data-v-954c30ab></line>`);
								});
								_push(`<!--]-->`);
							} else _push(`<!---->`);
							_push(`<!--]-->`);
						}
						_push(`<!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.ticks, (tick, index) => {
							_push(`<!--[-->`);
							if (tick.showLabel) _push(`<text${ssrRenderAttr("x", tick.labelX)}${ssrRenderAttr("y", tick.labelY)}${ssrRenderAttr("text-anchor", tick.labelAnchor)} style="${ssrRenderStyle({ fill: tick.color })}" class="chart-label chart-value-axis-label" data-v-954c30ab>${ssrInterpolate(tick.label)}</text>`);
							else _push(`<!---->`);
							_push(`<!--]-->`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.secondaryValueAxes, (axis) => {
							_push(`<!--[--><!--[-->`);
							ssrRenderList(axis.ticks, (tick, index) => {
								_push(`<!--[-->`);
								if (tick.showLabel) _push(`<text${ssrRenderAttr("data-axis-id", axis.id)}${ssrRenderAttr("x", tick.labelX)}${ssrRenderAttr("y", tick.labelY)}${ssrRenderAttr("text-anchor", tick.labelAnchor)} style="${ssrRenderStyle({ fill: tick.color })}" class="chart-label chart-value-axis-label chart-secondary-value-axis-label" data-v-954c30ab>${ssrInterpolate(tick.label)}</text>`);
								else _push(`<!---->`);
								_push(`<!--]-->`);
							});
							_push(`<!--]--><!--]-->`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.xTicks, (tick, index) => {
							_push(`<!--[-->`);
							if (tick.showLabel) _push(`<text${ssrRenderAttr("x", tick.labelX)}${ssrRenderAttr("y", tick.labelY)}${ssrRenderAttr("text-anchor", tick.labelAnchor)} style="${ssrRenderStyle({ fill: tick.color })}" class="chart-label chart-x-value-axis-label" data-v-954c30ab>${ssrInterpolate(tick.label)}</text>`);
							else _push(`<!---->`);
							_push(`<!--]-->`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.axisTicks, (tick, index) => {
							_push(`<line class="${ssrRenderClass(tick.minor ? "chart-axis-minor-tick" : "chart-axis-tick")}"${ssrRenderAttr("data-axis", tick.axis)}${ssrRenderAttr("x1", tick.x1)}${ssrRenderAttr("y1", tick.y1)}${ssrRenderAttr("x2", tick.x2)}${ssrRenderAttr("y2", tick.y2)} stroke="#64748b"${ssrRenderAttr("stroke-width", tick.minor ? 1 : 1.5)} data-v-954c30ab></line>`);
						});
						_push(`<!--]--><g${ssrRenderAttr("clip-path", `url(#chart-plot-${element.id})`)} data-v-954c30ab>`);
						if (element.chart.type !== "radar" && (["scatter", "bubble"].includes(element.chart.type) ? !element.chart.xAxis?.deleted : !element.chart.categoryAxisDeleted) && chartVisuals.value.get(element.id)?.baseline) _push(`<line class="${ssrRenderClass([{ "chart-x-axis-line": ["scatter", "bubble"].includes(element.chart.type) }, "chart-category-axis-line"])}"${ssrRenderAttr("data-axis", ["scatter", "bubble"].includes(element.chart.type) ? "xValue" : "category")}${ssrRenderAttr("x1", chartVisuals.value.get(element.id)?.baseline.x1)}${ssrRenderAttr("y1", chartVisuals.value.get(element.id)?.baseline.y1)}${ssrRenderAttr("x2", chartVisuals.value.get(element.id)?.baseline.x2)}${ssrRenderAttr("y2", chartVisuals.value.get(element.id)?.baseline.y2)} stroke="#64748b" stroke-width="2" data-v-954c30ab></line>`);
						else _push(`<!---->`);
						if (!element.chart.valueAxis?.deleted && chartVisuals.value.get(element.id)?.valueAxisLine) _push(`<line class="chart-value-axis-line"${ssrRenderAttr("x1", chartVisuals.value.get(element.id)?.valueAxisLine?.x1)}${ssrRenderAttr("y1", chartVisuals.value.get(element.id)?.valueAxisLine?.y1)}${ssrRenderAttr("x2", chartVisuals.value.get(element.id)?.valueAxisLine?.x2)}${ssrRenderAttr("y2", chartVisuals.value.get(element.id)?.valueAxisLine?.y2)} stroke="#64748b" stroke-width="2" data-v-954c30ab></line>`);
						else _push(`<!---->`);
						_push(`<!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.secondaryValueAxes, (axis) => {
							_push(`<!--[-->`);
							if (!axis.deleted) _push(`<line class="chart-value-axis-line chart-secondary-value-axis-line"${ssrRenderAttr("data-axis-id", axis.id)}${ssrRenderAttr("x1", axis.line.x1)}${ssrRenderAttr("y1", axis.line.y1)}${ssrRenderAttr("x2", axis.line.x2)}${ssrRenderAttr("y2", axis.line.y2)} stroke="#64748b" stroke-width="2" data-v-954c30ab></line>`);
							else _push(`<!---->`);
							_push(`<!--]-->`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.areas, (area, index) => {
							_push(`<path class="chart-area"${ssrRenderAttr("d", area.d)}${ssrRenderAttr("fill", chartPaint(element.id, area.color))}${ssrRenderAttr("fill-opacity", area.opacity ?? 1)}${ssrRenderAttr("stroke", chartPaint(element.id, area.stroke))}${ssrRenderAttr("stroke-width", area.strokeWidth ?? 1)}${ssrRenderAttr("stroke-dasharray", area.strokeDasharray)}${ssrRenderAttr("stroke-linecap", area.strokeLinecap)} data-v-954c30ab></path>`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.bars, (bar, index) => {
							_push(`<rect class="chart-bar"${ssrRenderAttr("x", bar.x)}${ssrRenderAttr("y", bar.y)}${ssrRenderAttr("width", bar.width)}${ssrRenderAttr("height", bar.height)}${ssrRenderAttr("fill", chartPaint(element.id, bar.color))}${ssrRenderAttr("stroke", chartPaint(element.id, bar.stroke))}${ssrRenderAttr("stroke-width", bar.strokeWidth ?? 1)}${ssrRenderAttr("stroke-dasharray", bar.strokeDasharray)}${ssrRenderAttr("stroke-linecap", bar.strokeLinecap)} data-v-954c30ab></rect>`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.lines, (line, index) => {
							_push(`<path class="chart-series-line"${ssrRenderAttr("d", line.d)} fill="none"${ssrRenderAttr("stroke", chartPaint(element.id, line.color))}${ssrRenderAttr("stroke-width", line.width ?? 4)}${ssrRenderAttr("stroke-dasharray", line.strokeDasharray)}${ssrRenderAttr("stroke-linecap", line.strokeLinecap ?? "round")} data-v-954c30ab></path>`);
						});
						_push(`<!--]--><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.markers, (marker, index) => {
							_push(`<!--[-->`);
							if (!chartMarkerPath(marker)) _push(`<circle class="${ssrRenderClass([{ "chart-bubble-marker": element.chart.type === "bubble" }, "chart-line-marker"])}"${ssrRenderAttr("data-marker-shape", marker.symbol ?? "circle")}${ssrRenderAttr("data-x", marker.x)}${ssrRenderAttr("data-y", marker.y)}${ssrRenderAttr("data-radius", marker.radius ?? 5)}${ssrRenderAttr("cx", marker.x)}${ssrRenderAttr("cy", marker.y)}${ssrRenderAttr("r", marker.symbol === "dot" ? (marker.radius ?? 5) * .45 : marker.radius ?? 5)}${ssrRenderAttr("fill", chartPaint(element.id, marker.color))}${ssrRenderAttr("stroke", chartPaint(element.id, marker.stroke))}${ssrRenderAttr("stroke-width", marker.strokeWidth ?? 1)}${ssrRenderAttr("stroke-dasharray", marker.strokeDasharray)}${ssrRenderAttr("stroke-linecap", marker.strokeLinecap)} data-v-954c30ab></circle>`);
							else _push(`<path class="chart-line-marker"${ssrRenderAttr("data-marker-shape", marker.symbol)}${ssrRenderAttr("data-x", marker.x)}${ssrRenderAttr("data-y", marker.y)}${ssrRenderAttr("data-radius", marker.radius ?? 5)}${ssrRenderAttr("d", chartMarkerPath(marker))}${ssrRenderAttr("fill", chartPaint(element.id, marker.color))}${ssrRenderAttr("stroke", chartPaint(element.id, marker.stroke))}${ssrRenderAttr("stroke-width", marker.strokeWidth ?? 1)}${ssrRenderAttr("stroke-dasharray", marker.strokeDasharray)}${ssrRenderAttr("stroke-linecap", marker.strokeLinecap)} data-v-954c30ab></path>`);
							_push(`<!--]-->`);
						});
						_push(`<!--]--></g><!--[-->`);
						ssrRenderList(chartVisuals.value.get(element.id)?.categories, (label, index) => {
							_push(`<text${ssrRenderAttr("x", label.x)}${ssrRenderAttr("y", label.y)}${ssrRenderAttr("text-anchor", label.anchor)} style="${ssrRenderStyle({ fill: label.color })}" class="chart-label chart-category-label" data-v-954c30ab>${ssrInterpolate(label.text)}</text>`);
						});
						_push(`<!--]--></g>`);
					}
					_push(`<!--[-->`);
					ssrRenderList(chartVisuals.value.get(element.id)?.dataLabels, (label, index) => {
						_push(`<g data-v-954c30ab>`);
						if (label.leaderLine) _push(`<path${ssrRenderAttr("d", label.leaderLine.d)}${ssrRenderAttr("stroke", label.leaderLine.color)} fill="none" stroke-width="1" class="chart-data-label-leader-line" data-v-954c30ab></path>`);
						else _push(`<!---->`);
						if (label.legendKey) {
							_push(`<g${ssrRenderAttr("data-key-id", chartDataLabelKeyId(element.id, index))}${ssrRenderAttr("fill", chartPaint(element.id, label.legendKey.color))} class="chart-data-label-key" data-v-954c30ab>`);
							if (!label.legendKey.markerChart) _push(`<rect${ssrRenderAttr("x", chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18)}${ssrRenderAttr("y", label.y - 12)} width="12" height="12" class="chart-data-label-key-swatch" data-v-954c30ab></rect>`);
							else _push(`<!---->`);
							if (label.legendKey.line) _push(`<path${ssrRenderAttr("d", `M ${chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18} ${label.y - 6} H ${(chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 12}`)} fill="none"${ssrRenderAttr("stroke", chartPaint(element.id, label.legendKey.color))} stroke-width="2" stroke-linecap="round" class="chart-data-label-key-line" data-v-954c30ab></path>`);
							else _push(`<!---->`);
							if (label.legendKey.symbol) {
								_push(`<!--[-->`);
								if (!chartMarkerPath({
									x: (chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 6,
									y: label.y - 6,
									color: label.legendKey.color,
									radius: 4,
									symbol: label.legendKey.symbol
								})) _push(`<circle${ssrRenderAttr("cx", (chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 6)}${ssrRenderAttr("cy", label.y - 6)}${ssrRenderAttr("r", label.legendKey.symbol === "dot" ? 1.8 : 4)} class="chart-data-label-key-marker"${ssrRenderAttr("data-marker-shape", label.legendKey.symbol)} data-v-954c30ab></circle>`);
								else _push(`<path${ssrRenderAttr("d", chartMarkerPath({
									x: (chartDataLabelKeyPositions.value.get(chartDataLabelKeyId(element.id, index)) ?? label.x - 18) + 6,
									y: label.y - 6,
									color: label.legendKey.color,
									radius: 4,
									symbol: label.legendKey.symbol
								}))} class="chart-data-label-key-marker"${ssrRenderAttr("data-marker-shape", label.legendKey.symbol)} data-v-954c30ab></path>`);
								_push(`<!--]-->`);
							} else _push(`<!---->`);
							_push(`</g>`);
						} else _push(`<!---->`);
						_push(`<text${ssrRenderAttr("x", label.x)}${ssrRenderAttr("y", label.y)}${ssrRenderAttr("text-anchor", label.anchor)} style="${ssrRenderStyle({
							fill: label.style?.color,
							fontFamily: label.style?.fontFamily ? unref(cssFontFamily)(label.style.fontFamily) : void 0,
							fontSize: label.style?.fontSizePt === void 0 ? void 0 : `${label.style.fontSizePt * 4 / 3}px`,
							fontWeight: label.style?.bold === void 0 ? void 0 : label.style.bold ? 700 : 400,
							fontStyle: label.style?.italic === void 0 ? void 0 : label.style.italic ? "italic" : "normal",
							letterSpacing: label.style?.letterSpacingPt === void 0 ? void 0 : `${label.style.letterSpacingPt * 4 / 3}px`,
							textDecorationLine: label.style?.underline === void 0 && label.style?.lineThrough === void 0 ? void 0 : [label.style?.underline ? "underline" : void 0, label.style?.lineThrough ? "line-through" : void 0].filter(Boolean).join(" ") || "none",
							textDecorationStyle: label.style?.underline ? label.style.underlineStyle : label.style?.lineThrough ? label.style.lineThroughStyle : void 0,
							textDecorationColor: label.style?.underline ? label.style.underlineColor : void 0,
							textDecorationThickness: label.style?.underline && label.style.underlineThicknessEmu !== void 0 ? `${label.style.underlineThicknessEmu * 4 / (3 * EMU_PER_POINT)}px` : void 0
						})}" class="chart-data-label" data-v-954c30ab><!--[-->`);
						ssrRenderList(label.lines, (line, lineIndex) => {
							_push(`<tspan${ssrRenderAttr("x", label.x)}${ssrRenderAttr("y", label.y + (lineIndex - (label.lines.length - 1) / 2) * 22)}${ssrRenderAttr("dy", label.style?.baselineOffsetEm === void 0 ? void 0 : `${-label.style.baselineOffsetEm}em`)} data-v-954c30ab>${ssrInterpolate(line)}</tspan>`);
						});
						_push(`<!--]--></text></g>`);
					});
					_push(`<!--]--><g class="chart-legend-group"${ssrRenderAttr("clip-path", chartVisuals.value.get(element.id)?.legendClip ? `url(#${chartVisuals.value.get(element.id)?.legendClip?.id})` : void 0)} data-v-954c30ab><!--[-->`);
					ssrRenderList(chartVisuals.value.get(element.id)?.legend, (entry, index) => {
						_push(`<g data-v-954c30ab><g${ssrRenderAttr("fill", chartPaint(element.id, entry.color))} class="chart-legend-key"${ssrRenderAttr("data-symbol", entry.symbol)} data-v-954c30ab>`);
						if (!entry.markerChart) _push(`<rect${ssrRenderAttr("x", entry.x)}${ssrRenderAttr("y", entry.y - 12)} width="12" height="12" data-v-954c30ab></rect>`);
						else _push(`<!---->`);
						if (entry.line) _push(`<path${ssrRenderAttr("d", `M ${entry.x} ${entry.y - 6} H ${entry.x + 12}`)} fill="none"${ssrRenderAttr("stroke", chartPaint(element.id, entry.color))} stroke-width="2" stroke-linecap="round" class="chart-legend-key-line" data-v-954c30ab></path>`);
						else _push(`<!---->`);
						if (entry.symbol) {
							_push(`<!--[-->`);
							if (!chartMarkerPath({
								x: entry.x + 6,
								y: entry.y - 6,
								color: entry.color,
								radius: 4,
								symbol: entry.symbol
							})) _push(`<circle${ssrRenderAttr("cx", entry.x + 6)}${ssrRenderAttr("cy", entry.y - 6)}${ssrRenderAttr("r", entry.symbol === "dot" ? 1.8 : 4)} class="chart-legend-key-marker"${ssrRenderAttr("data-marker-shape", entry.symbol)} data-v-954c30ab></circle>`);
							else _push(`<path${ssrRenderAttr("d", chartMarkerPath({
								x: entry.x + 6,
								y: entry.y - 6,
								color: entry.color,
								radius: 4,
								symbol: entry.symbol
							}))} class="chart-legend-key-marker"${ssrRenderAttr("data-marker-shape", entry.symbol)} data-v-954c30ab></path>`);
							_push(`<!--]-->`);
						} else _push(`<!---->`);
						_push(`</g><text${ssrRenderAttr("x", entry.x + 18)}${ssrRenderAttr("y", entry.y)} class="chart-label chart-legend-label" data-v-954c30ab>${ssrInterpolate(entry.text)}</text></g>`);
					});
					_push(`<!--]--></g></svg>`);
				} else if ((element.kind === "picture" || element.kind === "video") && element.imageUrl) {
					_push(`<div class="image-viewport" aria-hidden="true" data-v-954c30ab>`);
					if (element.kind === "picture" && (element.imageAlphaInverse !== void 0 || element.imageEffectDagAlphaInverse !== void 0)) {
						_push(`<svg class="image-tone-defs image-alpha-inverse-defs" aria-hidden="true" width="0" height="0" data-v-954c30ab><defs data-v-954c30ab><!--[-->`);
						ssrRenderList([element.imageAlphaInverse, element.imageEffectDagAlphaInverse], (alphaInverse, index) => {
							_push(`<!--[-->`);
							if (alphaInverse !== void 0) {
								_push(`<filter${ssrRenderAttr("id", imageAlphaInverseFilterId(element, index === 1))} class="image-alpha-inverse-filter" color-interpolation-filters="sRGB" data-v-954c30ab><feComponentTransfer data-v-954c30ab>`);
								if (alphaInverse) _push(`<feFuncR type="linear" slope="0"${ssrRenderAttr("intercept", alphaInverse[0])} data-v-954c30ab></feFuncR>`);
								else _push(`<!---->`);
								if (alphaInverse) _push(`<feFuncG type="linear" slope="0"${ssrRenderAttr("intercept", alphaInverse[1])} data-v-954c30ab></feFuncG>`);
								else _push(`<!---->`);
								if (alphaInverse) _push(`<feFuncB type="linear" slope="0"${ssrRenderAttr("intercept", alphaInverse[2])} data-v-954c30ab></feFuncB>`);
								else _push(`<!---->`);
								_push(`<feFuncA type="table" tableValues="1 0" data-v-954c30ab></feFuncA></feComponentTransfer></filter>`);
							} else _push(`<!---->`);
							_push(`<!--]-->`);
						});
						_push(`<!--]--></defs></svg>`);
					} else _push(`<!---->`);
					_push(`<!--[-->`);
					ssrRenderList(imageToneEnabled(element) ? [imageTone(element)] : [], (tone) => {
						_push(`<svg class="image-tone-defs" aria-hidden="true" width="0" height="0" data-v-954c30ab><defs data-v-954c30ab><filter${ssrRenderAttr("id", imageToneFilterId(element))} class="image-tone-filter" color-interpolation-filters="sRGB" data-v-954c30ab><feComponentTransfer data-v-954c30ab><feFuncR class="image-tone-gain" type="linear"${ssrRenderAttr("slope", tone.gain)}${ssrRenderAttr("intercept", (1 - tone.gain) / 2)} data-v-954c30ab></feFuncR><feFuncG class="image-tone-gain" type="linear"${ssrRenderAttr("slope", tone.gain)}${ssrRenderAttr("intercept", (1 - tone.gain) / 2)} data-v-954c30ab></feFuncG><feFuncB class="image-tone-gain" type="linear"${ssrRenderAttr("slope", tone.gain)}${ssrRenderAttr("intercept", (1 - tone.gain) / 2)} data-v-954c30ab></feFuncB></feComponentTransfer><feComponentTransfer data-v-954c30ab><feFuncR class="image-tone-blacklevel" type="linear" slope="1"${ssrRenderAttr("intercept", tone.blacklevel)} data-v-954c30ab></feFuncR><feFuncG class="image-tone-blacklevel" type="linear" slope="1"${ssrRenderAttr("intercept", tone.blacklevel)} data-v-954c30ab></feFuncG><feFuncB class="image-tone-blacklevel" type="linear" slope="1"${ssrRenderAttr("intercept", tone.blacklevel)} data-v-954c30ab></feFuncB></feComponentTransfer><feComponentTransfer data-v-954c30ab><feFuncR class="image-tone-gamma" type="gamma" amplitude="1"${ssrRenderAttr("exponent", 1 / tone.gamma)} offset="0" data-v-954c30ab></feFuncR><feFuncG class="image-tone-gamma" type="gamma" amplitude="1"${ssrRenderAttr("exponent", 1 / tone.gamma)} offset="0" data-v-954c30ab></feFuncG><feFuncB class="image-tone-gamma" type="gamma" amplitude="1"${ssrRenderAttr("exponent", 1 / tone.gamma)} offset="0" data-v-954c30ab></feFuncB></feComponentTransfer></filter></defs></svg>`);
					});
					_push(`<!--]--><img class="slide-image"${ssrRenderAttr("src", element.imageUrl)}${ssrRenderAttr("alt", element.name)} style="${ssrRenderStyle(imageStyle(element))}" data-v-954c30ab>`);
					if (pictureInnerShadowStyle(element)) _push(`<div class="picture-inner-shadow" style="${ssrRenderStyle(pictureInnerShadowStyle(element))}" aria-hidden="true" data-v-954c30ab></div>`);
					else _push(`<!---->`);
					_push(`</div>`);
				} else if (element.kind === "table" && element.table) {
					_push(`<table class="slide-table" data-v-954c30ab><colgroup data-v-954c30ab><!--[-->`);
					ssrRenderList(element.table.columns, (column, columnIndex) => {
						_push(`<col style="${ssrRenderStyle({ width: `${column / Math.max(element.table.columns.reduce((sum, width) => sum + width, 0), 1) * 100}%` })}" data-v-954c30ab>`);
					});
					_push(`<!--]--></colgroup><tbody data-v-954c30ab><!--[-->`);
					ssrRenderList(element.table.rows, (row, rowIndex) => {
						_push(`<tr style="${ssrRenderStyle({ height: `${row.height / Math.max(element.table.rows.reduce((sum, item) => sum + item.height, 0), 1) * 100}%` })}" data-v-954c30ab><!--[-->`);
						ssrRenderList(row.cells, (cell, cellIndex) => {
							_push(`<!--[-->`);
							if (!cell.hidden) {
								_push(`<td${ssrRenderAttr("colspan", cell.colSpan)}${ssrRenderAttr("rowspan", cell.rowSpan)}${ssrRenderAttr("data-autofit", cell.textAutoFitDynamic ? "normal" : void 0)}${ssrRenderAttr("data-row-index", rowIndex)}${ssrRenderAttr("data-cell-index", cellIndex)} style="${ssrRenderStyle(tableCellStyle(cell))}" data-v-954c30ab><div class="table-cell-text" style="${ssrRenderStyle(tableCellTextStyle(cell, element.table.rows.slice(rowIndex, rowIndex + cell.rowSpan).reduce((height, item) => height + item.height, 0)))}" data-v-954c30ab><!--[-->`);
								ssrRenderList(unref(paragraphEntries)(cell.paragraphs), (entry, paragraphIndex) => {
									_push(`<div class="text-paragraph" style="${ssrRenderStyle(paragraphStyle(entry.paragraph, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), cell.textLineSpacingReduction, cell.textColumnsRightToLeft && (cell.textColumnCount ?? 1) > 1 ? "ltr" : void 0))}" data-v-954c30ab>`);
									if (entry.paragraph.bullet?.kind === "image" && entry.paragraph.bullet.url) _push(`<img class="paragraph-picture-bullet"${ssrRenderAttr("src", entry.paragraph.bullet.url)} style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" alt="" aria-hidden="true" data-v-954c30ab>`);
									else if (entry.marker) _push(`<span class="paragraph-marker" style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" aria-hidden="true" data-v-954c30ab>${ssrInterpolate(entry.marker)}</span>`);
									else _push(`<!---->`);
									_push(`<!--[-->`);
									ssrRenderList(entry.paragraph.runs, (run, runIndex) => {
										_push(`<!--[--><!--[-->`);
										ssrRenderList(runSegments(run, element, paragraphIndex, runIndex), (segment, segmentIndex) => {
											ssrRenderVNode(_push, createVNode(resolveDynamicComponent(canActivateHyperlink(run.hyperlink) ? "a" : "span"), {
												key: segmentIndex,
												href: canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : void 0,
												title: run.hyperlink?.tooltip,
												"data-pptx-hyperlink": canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : void 0,
												class: ["text-run", {
													"text-run-hyperlink": canActivateHyperlink(run.hyperlink),
													"text-run-break": segment.text === "\n",
													"text-run-fill": !segment.separateDecorations && runHasTextFill(element, run, paragraphIndex, segment.start, segment.end),
													"text-run-highlight": Boolean(run.highlightColor)
												}],
												"data-char-start": segment.start,
												"data-char-end": segment.end,
												style: [
													runStyle(element, run, segment.script, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), paragraphIndex, segment.start, segment.end, segment.separateDecorations, "table-" + rowIndex + "-" + cellIndex + "-" + runIndex),
													textColorAnimationStyle(element, paragraphIndex, segment.start, segment.end),
													textRangeOpacityAnimationStyle(element, paragraphIndex, segment.start, segment.end),
													characterAnimationStyle(element, paragraphIndex, segment.start, segment.end),
													textRunAnimationStyle(element, paragraphIndex, segment.start, segment.end, run, effectiveTableCellFontScale(element, rowIndex, cellIndex, cell), segment.script, segment.separateDecorations, "table-" + rowIndex + "-" + cellIndex + "-" + runIndex),
													textTransformAnimationStyle(element, paragraphIndex, segment.start, segment.end)
												]
											}, {
												default: withCtx((_, _push, _parent, _scopeId) => {
													if (_push) {
														if (textInnerShadows(run).length) {
															_push(`<svg class="text-run-filter-defs" width="0" height="0" aria-hidden="true" focusable="false" data-v-954c30ab${_scopeId}><defs data-v-954c30ab${_scopeId}><filter${ssrRenderAttr("id", textInnerShadowFilterId(element, paragraphIndex, segment.start, segment.end, "table-" + rowIndex + "-" + cellIndex + "-" + runIndex))} x="-10%" y="-10%" width="120%" height="120%" filterUnits="objectBoundingBox" primitiveUnits="userSpaceOnUse" color-interpolation-filters="sRGB" data-v-954c30ab${_scopeId}><!--[-->`);
															ssrRenderList(textInnerShadows(run), (shadow, shadowIndex) => {
																_push(`<!--[--><feGaussianBlur in="SourceAlpha"${ssrRenderAttr("stdDeviation", Math.max(0, shadow.blur * scale.value / 2))}${ssrRenderAttr("result", "inner-blur-" + shadowIndex)} data-v-954c30ab${_scopeId}></feGaussianBlur><feOffset${ssrRenderAttr("in", "inner-blur-" + shadowIndex)}${ssrRenderAttr("dx", textInnerShadowOffset(shadow).dx)}${ssrRenderAttr("dy", textInnerShadowOffset(shadow).dy)}${ssrRenderAttr("result", "inner-offset-" + shadowIndex)} data-v-954c30ab${_scopeId}></feOffset><feComposite in="SourceAlpha"${ssrRenderAttr("in2", "inner-offset-" + shadowIndex)} operator="out"${ssrRenderAttr("result", "inner-mask-" + shadowIndex)} data-v-954c30ab${_scopeId}></feComposite><feFlood${ssrRenderAttr("flood-color", textInnerShadowColor(element, shadow, paragraphIndex, segment.start, segment.end))}${ssrRenderAttr("result", "inner-color-" + shadowIndex)} data-v-954c30ab${_scopeId}></feFlood><feComposite${ssrRenderAttr("in", "inner-color-" + shadowIndex)}${ssrRenderAttr("in2", "inner-mask-" + shadowIndex)} operator="in"${ssrRenderAttr("result", "inner-shadow-" + shadowIndex)} data-v-954c30ab${_scopeId}></feComposite><!--]-->`);
															});
															_push(`<!--]--><feMerge data-v-954c30ab${_scopeId}><!--[-->`);
															ssrRenderList(textInnerShadows(run), (_, shadowIndex) => {
																_push(`<feMergeNode${ssrRenderAttr("in", "inner-shadow-" + shadowIndex)} data-v-954c30ab${_scopeId}></feMergeNode>`);
															});
															_push(`<!--]--><feMergeNode in="SourceGraphic" data-v-954c30ab${_scopeId}></feMergeNode></feMerge></filter></defs></svg>`);
														} else _push(`<!---->`);
														if (segment.separateDecorations) _push(`<span class="${ssrRenderClass([{
															"text-run-fill": runHasTextFill(element, run, paragraphIndex, segment.start, segment.end),
															"text-run-highlight": Boolean(run.highlightColor)
														}, "text-run-underline"])}" style="${ssrRenderStyle(textUnderlineStyle(run))}" data-v-954c30ab${_scopeId}>${ssrInterpolate(segment.text)}</span>`);
														else _push(`<!--[-->${ssrInterpolate(segment.text)}<!--]-->`);
													} else return [textInnerShadows(run).length ? (openBlock(), createBlock("svg", {
														key: 0,
														class: "text-run-filter-defs",
														width: "0",
														height: "0",
														"aria-hidden": "true",
														focusable: "false"
													}, [createVNode("defs", null, [createVNode("filter", {
														id: textInnerShadowFilterId(element, paragraphIndex, segment.start, segment.end, "table-" + rowIndex + "-" + cellIndex + "-" + runIndex),
														x: "-10%",
														y: "-10%",
														width: "120%",
														height: "120%",
														filterUnits: "objectBoundingBox",
														primitiveUnits: "userSpaceOnUse",
														"color-interpolation-filters": "sRGB"
													}, [(openBlock(true), createBlock(Fragment, null, renderList(textInnerShadows(run), (shadow, shadowIndex) => {
														return openBlock(), createBlock(Fragment, { key: shadowIndex }, [
															createVNode("feGaussianBlur", {
																in: "SourceAlpha",
																stdDeviation: Math.max(0, shadow.blur * scale.value / 2),
																result: "inner-blur-" + shadowIndex
															}, null, 8, ["stdDeviation", "result"]),
															createVNode("feOffset", {
																in: "inner-blur-" + shadowIndex,
																dx: textInnerShadowOffset(shadow).dx,
																dy: textInnerShadowOffset(shadow).dy,
																result: "inner-offset-" + shadowIndex
															}, null, 8, [
																"in",
																"dx",
																"dy",
																"result"
															]),
															createVNode("feComposite", {
																in: "SourceAlpha",
																in2: "inner-offset-" + shadowIndex,
																operator: "out",
																result: "inner-mask-" + shadowIndex
															}, null, 8, ["in2", "result"]),
															createVNode("feFlood", {
																"flood-color": textInnerShadowColor(element, shadow, paragraphIndex, segment.start, segment.end),
																result: "inner-color-" + shadowIndex
															}, null, 8, ["flood-color", "result"]),
															createVNode("feComposite", {
																in: "inner-color-" + shadowIndex,
																in2: "inner-mask-" + shadowIndex,
																operator: "in",
																result: "inner-shadow-" + shadowIndex
															}, null, 8, [
																"in",
																"in2",
																"result"
															])
														], 64);
													}), 128)), createVNode("feMerge", null, [(openBlock(true), createBlock(Fragment, null, renderList(textInnerShadows(run), (_, shadowIndex) => {
														return openBlock(), createBlock("feMergeNode", {
															key: shadowIndex,
															in: "inner-shadow-" + shadowIndex
														}, null, 8, ["in"]);
													}), 128)), createVNode("feMergeNode", { in: "SourceGraphic" })])], 8, ["id"])])])) : createCommentVNode("v-if", true), segment.separateDecorations ? (openBlock(), createBlock("span", {
														key: 1,
														class: ["text-run-underline", {
															"text-run-fill": runHasTextFill(element, run, paragraphIndex, segment.start, segment.end),
															"text-run-highlight": Boolean(run.highlightColor)
														}],
														style: textUnderlineStyle(run)
													}, toDisplayString(segment.text), 7)) : (openBlock(), createBlock(Fragment, { key: 2 }, [createTextVNode(toDisplayString(segment.text), 1)], 64))];
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
				} else if (element.kind === "placeholder") _push(`<span data-v-954c30ab>${ssrInterpolate(element.mediaType ? "Media unavailable" : "Image unavailable")}</span>`);
				else _push(`<!---->`);
				if (element.paragraphs.length) {
					_push(`<div class="text-frame"${ssrRenderAttr("data-element-id", element.id)}${ssrRenderAttr("data-autofit", element.textAutoFit === "shape" ? "shape" : element.textAutoFitDynamic ? "normal" : void 0)}${ssrRenderAttr("data-text-overflow", supportsTextVerticalEllipsis(element) ? "ellipsis" : void 0)} style="${ssrRenderStyle(textFrameStyle(element))}" data-v-954c30ab><div class="text-bounds" style="${ssrRenderStyle(textBoundsStyle(element))}" data-v-954c30ab><!--[-->`);
					ssrRenderList(unref(paragraphEntries)(element.paragraphs), (entry, index) => {
						_push(`<div class="text-paragraph" style="${ssrRenderStyle([paragraphStyle(entry.paragraph, effectiveTextFontScale(element), element.textLineSpacingReduction, usesRightToLeftColumns(element) ? "ltr" : void 0), paragraphAnimationStyle(element, index)])}" data-v-954c30ab>`);
						if (entry.paragraph.bullet?.kind === "image" && entry.paragraph.bullet.url) _push(`<img class="paragraph-picture-bullet"${ssrRenderAttr("src", entry.paragraph.bullet.url)} style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" alt="" aria-hidden="true" data-v-954c30ab>`);
						else if (entry.marker) _push(`<span class="paragraph-marker" style="${ssrRenderStyle(paragraphBulletStyle(entry.paragraph))}" aria-hidden="true" data-v-954c30ab>${ssrInterpolate(entry.marker)}</span>`);
						else _push(`<!---->`);
						_push(`<!--[-->`);
						ssrRenderList(entry.paragraph.runs, (run, runIndex) => {
							_push(`<!--[--><!--[-->`);
							ssrRenderList(runSegments(run, element, index, runIndex), (segment, segmentIndex) => {
								ssrRenderVNode(_push, createVNode(resolveDynamicComponent(canActivateHyperlink(run.hyperlink) ? "a" : "span"), {
									key: segmentIndex,
									href: canActivateHyperlink(run.hyperlink) ? hyperlinkHref(run.hyperlink) : void 0,
									title: run.hyperlink?.tooltip,
									"data-pptx-hyperlink": canActivateHyperlink(run.hyperlink) ? JSON.stringify(run.hyperlink) : void 0,
									class: ["text-run", {
										"text-run-hyperlink": canActivateHyperlink(run.hyperlink),
										"text-run-break": segment.text === "\n",
										"text-run-fill": !segment.separateDecorations && runHasTextFill(element, run, index, segment.start, segment.end),
										"text-run-highlight": Boolean(run.highlightColor)
									}],
									"data-char-start": segment.start,
									"data-char-end": segment.end,
									style: [
										runStyle(element, run, segment.script, effectiveTextFontScale(element), index, segment.start, segment.end, segment.separateDecorations, "text-" + runIndex),
										textColorAnimationStyle(element, index, segment.start, segment.end),
										textRangeOpacityAnimationStyle(element, index, segment.start, segment.end),
										characterAnimationStyle(element, index, segment.start, segment.end),
										textRunAnimationStyle(element, index, segment.start, segment.end, run, effectiveTextFontScale(element), segment.script, segment.separateDecorations, "text-" + runIndex),
										textTransformAnimationStyle(element, index, segment.start, segment.end)
									]
								}, {
									default: withCtx((_, _push, _parent, _scopeId) => {
										if (_push) {
											if (textInnerShadows(run).length) {
												_push(`<svg class="text-run-filter-defs" width="0" height="0" aria-hidden="true" focusable="false" data-v-954c30ab${_scopeId}><defs data-v-954c30ab${_scopeId}><filter${ssrRenderAttr("id", textInnerShadowFilterId(element, index, segment.start, segment.end, "text-" + runIndex))} x="-10%" y="-10%" width="120%" height="120%" filterUnits="objectBoundingBox" primitiveUnits="userSpaceOnUse" color-interpolation-filters="sRGB" data-v-954c30ab${_scopeId}><!--[-->`);
												ssrRenderList(textInnerShadows(run), (shadow, shadowIndex) => {
													_push(`<!--[--><feGaussianBlur in="SourceAlpha"${ssrRenderAttr("stdDeviation", Math.max(0, shadow.blur * scale.value / 2))}${ssrRenderAttr("result", "inner-blur-" + shadowIndex)} data-v-954c30ab${_scopeId}></feGaussianBlur><feOffset${ssrRenderAttr("in", "inner-blur-" + shadowIndex)}${ssrRenderAttr("dx", textInnerShadowOffset(shadow).dx)}${ssrRenderAttr("dy", textInnerShadowOffset(shadow).dy)}${ssrRenderAttr("result", "inner-offset-" + shadowIndex)} data-v-954c30ab${_scopeId}></feOffset><feComposite in="SourceAlpha"${ssrRenderAttr("in2", "inner-offset-" + shadowIndex)} operator="out"${ssrRenderAttr("result", "inner-mask-" + shadowIndex)} data-v-954c30ab${_scopeId}></feComposite><feFlood${ssrRenderAttr("flood-color", textInnerShadowColor(element, shadow, index, segment.start, segment.end))}${ssrRenderAttr("result", "inner-color-" + shadowIndex)} data-v-954c30ab${_scopeId}></feFlood><feComposite${ssrRenderAttr("in", "inner-color-" + shadowIndex)}${ssrRenderAttr("in2", "inner-mask-" + shadowIndex)} operator="in"${ssrRenderAttr("result", "inner-shadow-" + shadowIndex)} data-v-954c30ab${_scopeId}></feComposite><!--]-->`);
												});
												_push(`<!--]--><feMerge data-v-954c30ab${_scopeId}><!--[-->`);
												ssrRenderList(textInnerShadows(run), (_, shadowIndex) => {
													_push(`<feMergeNode${ssrRenderAttr("in", "inner-shadow-" + shadowIndex)} data-v-954c30ab${_scopeId}></feMergeNode>`);
												});
												_push(`<!--]--><feMergeNode in="SourceGraphic" data-v-954c30ab${_scopeId}></feMergeNode></feMerge></filter></defs></svg>`);
											} else _push(`<!---->`);
											if (segment.separateDecorations) _push(`<span class="${ssrRenderClass([{
												"text-run-fill": runHasTextFill(element, run, index, segment.start, segment.end),
												"text-run-highlight": Boolean(run.highlightColor)
											}, "text-run-underline"])}" style="${ssrRenderStyle(textUnderlineStyle(run))}" data-v-954c30ab${_scopeId}>${ssrInterpolate(segment.text)}</span>`);
											else _push(`<!--[-->${ssrInterpolate(segment.text)}<!--]-->`);
										} else return [textInnerShadows(run).length ? (openBlock(), createBlock("svg", {
											key: 0,
											class: "text-run-filter-defs",
											width: "0",
											height: "0",
											"aria-hidden": "true",
											focusable: "false"
										}, [createVNode("defs", null, [createVNode("filter", {
											id: textInnerShadowFilterId(element, index, segment.start, segment.end, "text-" + runIndex),
											x: "-10%",
											y: "-10%",
											width: "120%",
											height: "120%",
											filterUnits: "objectBoundingBox",
											primitiveUnits: "userSpaceOnUse",
											"color-interpolation-filters": "sRGB"
										}, [(openBlock(true), createBlock(Fragment, null, renderList(textInnerShadows(run), (shadow, shadowIndex) => {
											return openBlock(), createBlock(Fragment, { key: shadowIndex }, [
												createVNode("feGaussianBlur", {
													in: "SourceAlpha",
													stdDeviation: Math.max(0, shadow.blur * scale.value / 2),
													result: "inner-blur-" + shadowIndex
												}, null, 8, ["stdDeviation", "result"]),
												createVNode("feOffset", {
													in: "inner-blur-" + shadowIndex,
													dx: textInnerShadowOffset(shadow).dx,
													dy: textInnerShadowOffset(shadow).dy,
													result: "inner-offset-" + shadowIndex
												}, null, 8, [
													"in",
													"dx",
													"dy",
													"result"
												]),
												createVNode("feComposite", {
													in: "SourceAlpha",
													in2: "inner-offset-" + shadowIndex,
													operator: "out",
													result: "inner-mask-" + shadowIndex
												}, null, 8, ["in2", "result"]),
												createVNode("feFlood", {
													"flood-color": textInnerShadowColor(element, shadow, index, segment.start, segment.end),
													result: "inner-color-" + shadowIndex
												}, null, 8, ["flood-color", "result"]),
												createVNode("feComposite", {
													in: "inner-color-" + shadowIndex,
													in2: "inner-mask-" + shadowIndex,
													operator: "in",
													result: "inner-shadow-" + shadowIndex
												}, null, 8, [
													"in",
													"in2",
													"result"
												])
											], 64);
										}), 128)), createVNode("feMerge", null, [(openBlock(true), createBlock(Fragment, null, renderList(textInnerShadows(run), (_, shadowIndex) => {
											return openBlock(), createBlock("feMergeNode", {
												key: shadowIndex,
												in: "inner-shadow-" + shadowIndex
											}, null, 8, ["in"]);
										}), 128)), createVNode("feMergeNode", { in: "SourceGraphic" })])], 8, ["id"])])])) : createCommentVNode("v-if", true), segment.separateDecorations ? (openBlock(), createBlock("span", {
											key: 1,
											class: ["text-run-underline", {
												"text-run-fill": runHasTextFill(element, run, index, segment.start, segment.end),
												"text-run-highlight": Boolean(run.highlightColor)
											}],
											style: textUnderlineStyle(run)
										}, toDisplayString(segment.text), 7)) : (openBlock(), createBlock(Fragment, { key: 2 }, [createTextVNode(toDisplayString(segment.text), 1)], 64))];
									}),
									_: 2
								}), _parent);
							});
							_push(`<!--]--><!--]-->`);
						});
						_push(`<!--]--></div>`);
					});
					_push(`<!--]--></div></div>`);
				} else _push(`<!---->`);
				if (element.kind === "line") {
					_push(`<svg class="line-render" style="${ssrRenderStyle(lineRenderStyle(element))}" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true" data-v-954c30ab><defs data-v-954c30ab>`);
					if (element.strokePattern) _push(`<pattern${ssrRenderAttr("id", customGeometryStrokePatternId(element))} patternUnits="userSpaceOnUse" width="1000" height="1000" data-v-954c30ab><foreignObject width="1000" height="1000" data-v-954c30ab><div xmlns="http://www.w3.org/1999/xhtml" style="${ssrRenderStyle([{
						"width": "100%",
						"height": "100%"
					}, { background: element.strokePattern }])}" data-v-954c30ab></div></foreignObject></pattern>`);
					else _push(`<!---->`);
					if (animationLineEnd(element, "head").type !== "none") _push(`<marker${ssrRenderAttr("id", lineEndId(element, "head"))}${ssrRenderAttr("markerWidth", lineEndMarkerSize(element, "head").width)}${ssrRenderAttr("markerHeight", lineEndMarkerSize(element, "head").height)} viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto-start-reverse" data-v-954c30ab><path${ssrRenderAttr("d", lineEndPath(animationLineEnd(element, "head").type))}${ssrRenderAttr("fill", lineEndUsesStroke(animationLineEnd(element, "head").type) ? "none" : lineStrokeColor(element))}${ssrRenderAttr("stroke", lineEndUsesStroke(animationLineEnd(element, "head").type) ? lineStrokeColor(element) : "none")} data-v-954c30ab></path></marker>`);
					else _push(`<!---->`);
					if (animationLineEnd(element, "tail").type !== "none") _push(`<marker${ssrRenderAttr("id", lineEndId(element, "tail"))}${ssrRenderAttr("markerWidth", lineEndMarkerSize(element, "tail").width)}${ssrRenderAttr("markerHeight", lineEndMarkerSize(element, "tail").height)} viewBox="0 0 10 10" refX="10" refY="5" markerUnits="strokeWidth" orient="auto" data-v-954c30ab><path${ssrRenderAttr("d", lineEndPath(animationLineEnd(element, "tail").type))}${ssrRenderAttr("fill", lineEndUsesStroke(animationLineEnd(element, "tail").type) ? "none" : lineStrokeColor(element))}${ssrRenderAttr("stroke", lineEndUsesStroke(animationLineEnd(element, "tail").type) ? lineStrokeColor(element) : "none")} data-v-954c30ab></path></marker>`);
					else _push(`<!---->`);
					_push(`</defs>`);
					if (element.linePath) _push(`<path${ssrRenderAttr("d", element.linePath)} fill="none"${ssrRenderAttr("stroke", lineStrokeColor(element))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}"${ssrRenderAttr("marker-start", lineEndUrl(element, "head"))}${ssrRenderAttr("marker-end", lineEndUrl(element, "tail"))} vector-effect="non-scaling-stroke" data-v-954c30ab></path>`);
					else _push(`<line x1="0"${ssrRenderAttr("y1", element.geometry === "lineInv" ? 1e3 : 0)}${ssrRenderAttr("x2", element.width > 0 ? 1e3 : 0)}${ssrRenderAttr("y2", element.height > 0 ? element.geometry === "lineInv" ? 0 : 1e3 : 0)}${ssrRenderAttr("stroke", lineStrokeColor(element))}${ssrRenderAttr("stroke-width", customGeometryStrokeWidth(element))} style="${ssrRenderStyle(customGeometryStrokeStyle(element))}"${ssrRenderAttr("marker-start", lineEndUrl(element, "head"))}${ssrRenderAttr("marker-end", lineEndUrl(element, "tail"))} vector-effect="non-scaling-stroke" data-v-954c30ab></line>`);
					_push(`</svg>`);
				} else _push(`<!---->`);
				_push(`</div>`);
			});
			_push(`<!--]-->`);
			if (backgroundTileWarning.value) _push(`<span class="media-warning" role="status" data-v-954c30ab>${ssrInterpolate(backgroundTileWarning.value)}</span>`);
			else _push(`<!---->`);
			if (textOverflowWarning.value && (!__props.thumbnail || __props.playbackPreview)) _push(`<span class="media-warning" role="status" data-v-954c30ab>${ssrInterpolate(textOverflowWarning.value)}</span>`);
			else _push(`<!---->`);
			_push(`</div>`);
		};
	}
});
//#endregion
//#region app/components/SlideStage.vue
var _sfc_setup = SlideStage_vue_vue_type_script_setup_true_lang_default.setup;
SlideStage_vue_vue_type_script_setup_true_lang_default.setup = (props, ctx) => {
	const ssrContext = useSSRContext();
	(ssrContext.modules || (ssrContext.modules = /* @__PURE__ */ new Set())).add("components/SlideStage.vue");
	return _sfc_setup ? _sfc_setup(props, ctx) : void 0;
};
var SlideStage_default = /*#__PURE__*/ Object.assign(_plugin_vue_export_helper_default(SlideStage_vue_vue_type_script_setup_true_lang_default, [["__scopeId", "data-v-954c30ab"]]), { __name: "SlideStage" });
//#endregion
//#region app/composables/useSlideTransition.ts
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
function withoutMorphTransforms(elements, measure) {
	const styles = /* @__PURE__ */ new Map();
	for (const element of elements) for (let current = element; current; current = current.parentElement) {
		if (getComputedStyle(current).transform !== "none" && !styles.has(current)) {
			styles.set(current, [current.style.getPropertyValue("transform"), current.style.getPropertyPriority("transform")]);
			current.style.setProperty("transform", "none", "important");
		}
		if (current.classList.contains("slide-frame")) break;
	}
	try {
		return measure();
	} finally {
		for (const [element, [value, priority]] of styles) if (value) element.style.setProperty("transform", value, priority);
		else element.style.removeProperty("transform");
	}
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
			const computed = getComputedStyle(start.node.parentElement || frame);
			tokens.push({
				text: segment.text,
				key: segment.key,
				left: rect.left - frameRect.left,
				top: rect.top - frameRect.top,
				width: rect.width,
				height: rect.height,
				content: range.cloneContents(),
				style: {
					color: computed.color,
					fontFamily: computed.fontFamily,
					fontSize: computed.fontSize,
					fontStyle: computed.fontStyle,
					fontWeight: computed.fontWeight,
					fontVariant: computed.fontVariant,
					fontStretch: computed.fontStretch,
					lineHeight: computed.lineHeight,
					letterSpacing: computed.letterSpacing,
					wordSpacing: computed.wordSpacing,
					textDecoration: computed.textDecoration,
					textShadow: computed.textShadow,
					textTransform: computed.textTransform,
					direction: computed.direction,
					writingMode: computed.writingMode,
					verticalAlign: computed.verticalAlign
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
		width: `${Math.max(token.width, .01)}px`,
		height: `${Math.max(token.height, .01)}px`,
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
		if (nextSlide?.id === previousSlide?.id) return;
		morphCleanup?.();
		morphCleanup = void 0;
		underlay.value?.replaceChildren();
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
		underlay.value?.append(snapshot);
	}, { flush: "sync" });
	watch(slide, async (activeSlide, previousSlide) => {
		if (!activeSlide || activeSlide.id === previousSlide?.id) return;
		const currentPlayback = playback.value?.slideId === activeSlide.id ? playback.value : null;
		const transition = currentPlayback?.transition;
		await nextTick();
		if (slide.value?.id !== activeSlide.id) return;
		const animationDelay = currentPlayback ? Math.min(0, currentPlayback.startedAtEpoch - (performance.timeOrigin + performance.now())) : 0;
		shell.value?.classList.remove("is-fading-through-black", "is-cut-through-black", "is-prism-transition", "is-flip-transition", "is-switch-transition", "is-ferris-transition", "is-gallery-transition", "is-conveyor-transition", "is-morph-transition", "is-preset-transition");
		underlay.value?.classList.remove("is-shredding-out", "is-flythrough", "is-warping-out", "is-vortexing");
		const incoming = incomingHost.value;
		incoming?.getAnimations().forEach((animation) => animation.cancel());
		flash.value?.getAnimations().forEach((animation) => animation.cancel());
		if (!transition || (void 0).matchMedia("(prefers-reduced-motion: reduce)").matches || !incoming || !underlay.value?.firstElementChild) {
			underlay.value?.replaceChildren();
			return;
		}
		if (transition.effect === "morph") {
			const outgoing = underlay.value?.firstElementChild;
			const stage = shell.value;
			if (!outgoing || !stage) {
				underlay.value?.replaceChildren();
				return;
			}
			const timing = {
				duration: transition.durationMs,
				easing: "ease-in-out",
				fill: "both",
				delay: animationDelay
			};
			const animations = [];
			const textMode = transition.morphOption === "byWord" || transition.morphOption === "byChar" ? transition.morphOption : void 0;
			stage.classList.add("is-morph-transition");
			const outgoingFrame = outgoing.querySelector(".slide-frame");
			const incomingFrame = incoming.querySelector(".slide-frame");
			const outgoingElements = [...outgoingFrame?.querySelectorAll(".slide-element") || []];
			const incomingElements = [...incomingFrame?.querySelectorAll(".slide-element") || []];
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
			if (textMode && incomingFrame) for (const [key, element] of newByName) {
				const previous = oldByName.get(key);
				if (!previous) continue;
				const sourceText = previous.querySelector(".text-frame");
				const targetText = element.querySelector(".text-frame");
				if (!sourceText || !targetText) continue;
				if (element.dataset.elementKind !== "shape") {
					console.warn(`Morph ${textMode} text matching is not supported on ${element.dataset.elementKind} elements; object motion is used.`);
					continue;
				}
				const [sourceTokens, targetTokens] = withoutMorphTransforms([previous, element], () => [morphTextTokens(sourceText, textMode), morphTextTokens(targetText, textMode)]);
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
						const scaleX = token.width / Math.max(target.width, .01);
						const scaleY = token.height / Math.max(target.height, .01);
						animations.push(node.animate([{
							left: `${token.left}px`,
							top: `${token.top}px`,
							transform: `scale(${scaleX}, ${scaleY})`
						}, {
							left: `${target.left}px`,
							top: `${target.top}px`,
							transform: "scale(1, 1)"
						}], timing));
					} else {
						const node = morphTextNode(layer, token, token.left, token.top);
						animations.push(node.animate([{ opacity: 1 }, { opacity: 0 }], timing));
					}
				}
				for (const token of targetTokens) {
					if (matchedTargets.has(token)) continue;
					const node = morphTextNode(layer, token, token.left, token.top);
					animations.push(node.animate([{ opacity: 0 }, { opacity: 1 }], timing));
				}
				for (const run of sourceText.querySelectorAll(".text-run")) run.style.setProperty("visibility", "hidden", "important");
				for (const run of targetText.querySelectorAll(".text-run")) {
					const visibility = run.style.getPropertyValue("visibility");
					const priority = run.style.getPropertyPriority("visibility");
					run.style.setProperty("visibility", "hidden", "important");
					restoreTextStyles.push(() => visibility ? run.style.setProperty("visibility", visibility, priority) : run.style.removeProperty("visibility"));
				}
			}
			for (const element of incomingElements) {
				const key = `${element.dataset.elementKind}:${element.dataset.elementName}`;
				const previous = oldByName.get(key);
				const finalOpacity = getComputedStyle(element).opacity;
				if (!previous || !newByName.has(key)) {
					animations.push(element.animate([{ opacity: 0 }, { opacity: finalOpacity }], timing));
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
				}], timing));
			}
			for (const element of outgoingElements) animations.push(element.animate([{ opacity: getComputedStyle(element).opacity }, { opacity: 0 }], timing));
			if (incomingFrame) {
				const backgroundAnimation = incomingFrame.animate([{ backgroundColor: "rgba(0, 0, 0, 0)" }, { backgroundColor: getComputedStyle(incomingFrame).backgroundColor }], timing);
				animations.push(backgroundAnimation);
				let cleaned = false;
				const cleanup = () => {
					if (cleaned) return;
					cleaned = true;
					animations.forEach((animation) => animation.cancel());
					restoreTextStyles.forEach((restore) => restore());
					incomingFrame.querySelectorAll(".morph-text-layer").forEach((layer) => layer.remove());
					if (morphCleanup === cleanup) morphCleanup = void 0;
				};
				morphCleanup = cleanup;
				backgroundAnimation.onfinish = () => {
					if (slide.value?.id === activeSlide.id) {
						cleanup();
						underlay.value?.replaceChildren();
						stage.classList.remove("is-morph-transition");
					}
				};
			}
			return;
		}
		if (transition.effect === "preset") {
			const outgoing = underlay.value?.firstElementChild;
			const stage = shell.value;
			if (!outgoing || !stage || !transition.presetName) {
				underlay.value?.replaceChildren();
				return;
			}
			const x = transition.invertX ? -1 : 1;
			const y = transition.invertY ? -1 : 1;
			const origin = x > 0 ? "left center" : "right center";
			const frames = (() => {
				switch (transition.presetName) {
					case "fallOver": return [{
						transform: "perspective(1200px) rotateX(0) rotateZ(0)",
						opacity: 1
					}, {
						transform: `perspective(1200px) rotateX(${72 * y}deg) rotateZ(${5 * x}deg) translateY(${10 * y}%)`,
						opacity: 0
					}];
					case "drape": return [
						{
							transform: "perspective(1200px) rotateX(0) scaleY(1)",
							opacity: 1
						},
						{
							transform: `perspective(1200px) rotateX(${48 * y}deg) scaleY(.9)`,
							opacity: .7,
							offset: .55
						},
						{
							transform: `perspective(1200px) rotateX(${78 * y}deg) scaleY(.7)`,
							opacity: 0
						}
					];
					case "curtains": return [{
						transform: "scaleX(1) rotateY(0)",
						transformOrigin: origin,
						opacity: 1
					}, {
						transform: `scaleX(.02) rotateY(${8 * y}deg)`,
						transformOrigin: origin,
						opacity: 0
					}];
					case "wind": return [{
						transform: "translate(0, 0) skewX(0)",
						opacity: 1
					}, {
						transform: `translate(${125 * x}%, ${ -110 * y}%) skewX(${14 * y}deg)`,
						opacity: 0
					}];
					case "prestige": return [{
						transform: "scale(1) rotate(0)",
						opacity: 1
					}, {
						transform: `scale(.62) rotate(${5 * x}deg) translate(${5 * x}%, ${12 * y}%)`,
						opacity: 0
					}];
					case "fracture": return [{
						clipPath: "inset(0)",
						transform: "scale(1) rotate(0)",
						opacity: 1
					}, {
						clipPath: `polygon(50% ${50 - 4 * y}%, ${50 + 4 * x}% 50%, 50% ${50 + 4 * y}%, ${50 - 4 * x}% 50%)`,
						transform: `scale(.82) rotate(${8 * x}deg)`,
						opacity: 0
					}];
					case "crush": return [{
						transform: "scale(1)",
						opacity: 1
					}, {
						transform: "scaleX(.04)",
						opacity: 0
					}];
					case "peelOff": return [{
						transform: "perspective(1200px) rotateY(0)",
						transformOrigin: origin,
						opacity: 1
					}, {
						transform: `perspective(1200px) rotateY(${105 * x}deg)`,
						transformOrigin: origin,
						opacity: 0
					}];
					case "pageCurlDouble": return [
						{
							transform: "perspective(1200px) rotateY(0) scaleX(1)",
							transformOrigin: "center center",
							opacity: 1
						},
						{
							transform: `perspective(1200px) rotateY(${82 * x}deg) scaleX(.28)`,
							transformOrigin: "center center",
							opacity: .8,
							offset: .55
						},
						{
							transform: `perspective(1200px) rotateY(${165 * x}deg) scaleX(.04)`,
							transformOrigin: "center center",
							opacity: 0
						}
					];
					case "pageCurlSingle": return [{
						transform: "perspective(1200px) rotateY(0)",
						transformOrigin: origin,
						opacity: 1
					}, {
						transform: `perspective(1200px) rotateY(${112 * x}deg)`,
						transformOrigin: origin,
						opacity: 0
					}];
					case "airplane": return [{
						transform: "perspective(1200px) translate(0, 0) rotate(0) rotateY(0) scale(1)",
						opacity: 1
					}, {
						transform: `perspective(1200px) translate(${130 * x}%, ${95 * y}%) rotate(${18 * x}deg) rotateY(${42 * y}deg) scale(.72)`,
						opacity: 0
					}];
					case "origami": return [
						{
							transform: "perspective(1200px) rotateX(0) rotateY(0) scale(1)",
							opacity: 1
						},
						{
							transform: `perspective(1200px) rotateX(${24 * y}deg) rotateY(${34 * x}deg) scale(.88)`,
							opacity: .9,
							offset: .32
						},
						{
							transform: `perspective(1200px) rotateX(${ -32 * y}deg) rotateY(${ -42 * x}deg) scale(.68)`,
							opacity: .65,
							offset: .68
						},
						{
							transform: `perspective(1200px) rotateX(${90 * y}deg) rotateY(${100 * x}deg) translate(${30 * x}%, ${30 * y}%) scale(.48)`,
							opacity: 0
						}
					];
				}
			})();
			stage.classList.add("is-preset-transition");
			const animation = outgoing.animate(frames, {
				duration: transition.durationMs,
				easing: "ease-in-out",
				fill: "both",
				delay: animationDelay
			});
			animation.onfinish = () => {
				animation.cancel();
				if (slide.value?.id === activeSlide.id) {
					underlay.value?.replaceChildren();
					stage.classList.remove("is-preset-transition");
				}
			};
			return;
		}
		if (transition.effect === "cut") {
			underlay.value?.replaceChildren();
			if (transition.throughBlack) {
				const stage = shell.value;
				stage?.classList.add("is-cut-through-black");
				requestAnimationFrame(() => requestAnimationFrame(() => {
					if (slide.value?.id === activeSlide.id) stage?.classList.remove("is-cut-through-black");
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
			const outgoing = underlay.value?.firstElementChild;
			const stage = shell.value;
			if (!outgoing || !stage) {
				underlay.value?.replaceChildren();
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
			const outgoingBase = outgoing.style.transform || "none";
			const base = (transform) => transform === "none" ? "" : `${transform} `;
			const side = direction === "r" ? -1 : 1;
			const startTransform = effect === "ferris" ? `${base(incomingBase)}perspective(1200px) rotateZ(${side * 180}deg) scale(0.78)` : effect === "gallery" ? `${base(incomingBase)}translateX(${side * 36}%) perspective(1200px) rotateY(${side * 38}deg) scale(0.84)` : effect === "conveyor" ? `${base(incomingBase)}translateX(${side * 100}%) perspective(1200px) rotateX(-10deg) rotateY(${side * 12}deg)` : `${base(incomingBase)}rotate3d(${rotation.axis}, ${angle}deg)`;
			const endTransform = effect === "ferris" ? `${base(outgoingBase)}perspective(1200px) rotateZ(${side * -180}deg) scale(0.78)` : effect === "gallery" ? `${base(outgoingBase)}translateX(${side * -36}%) perspective(1200px) rotateY(${side * -38}deg) scale(0.84)` : effect === "conveyor" ? `${base(outgoingBase)}translateX(${side * -100}%) perspective(1200px) rotateX(10deg) rotateY(${side * -12}deg)` : `${base(outgoingBase)}rotate3d(${rotation.axis}, ${-angle}deg)`;
			const transitionClass = effect === "prism" ? "is-prism-transition" : effect === "flip" ? "is-flip-transition" : effect === "switch" ? "is-switch-transition" : `is-${effect}-transition`;
			stage.classList.add(transitionClass);
			const timing = {
				duration: transition.durationMs,
				easing: "ease-in-out",
				fill: "both",
				delay: animationDelay
			};
			const incomingAnimation = incoming.animate([{ transform: startTransform }, { transform: incomingBase }], timing);
			const outgoingAnimation = outgoing.animate([{ transform: outgoingBase }, { transform: endTransform }], timing);
			incomingAnimation.onfinish = () => {
				incomingAnimation.cancel();
				outgoingAnimation.cancel();
				if (slide.value?.id === activeSlide.id) {
					underlay.value?.replaceChildren();
					stage.classList.remove(transitionClass);
				}
			};
			return;
		}
		if (effect === "flash") {
			const animation = flash.value?.animate([
				{ opacity: 0 },
				{
					opacity: 1,
					offset: .12
				},
				{
					opacity: 1,
					offset: .28
				},
				{ opacity: 0 }
			], {
				duration: transition.durationMs,
				easing: "linear",
				fill: "both",
				delay: animationDelay
			});
			if (animation) animation.onfinish = () => animation.cancel();
			underlay.value?.replaceChildren();
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
				offset: .5
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
					offset: .5
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
				.25,
				.5,
				.75,
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
				.25,
				.5,
				.75,
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
				.25,
				.5,
				.75,
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
				.25,
				.5,
				.75,
				1
			] : [
				0,
				.5,
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
				.25,
				.5,
				.75,
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
				.25,
				.5,
				.75,
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
				.25,
				.5,
				.75,
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
				.2,
				.4,
				.6,
				.8,
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
				.25,
				.5,
				.75,
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
				.25,
				.5,
				.75,
				1
			];
			shredOutgoing = transition.direction === "out";
			if (shredOutgoing) underlay.value?.classList.add("is-shredding-out");
			frames = progressValues.map((progress, index) => ({
				clipPath: shredTransitionClipPath(width, height, transition.shredPattern || "strip", shredOutgoing ? 1 - progress : progress),
				offset: index / (progressValues.length - 1)
			}));
		} else if (effect === "flythrough") {
			const baseTransform = incoming.style.transform === "none" ? "" : `${incoming.style.transform} `;
			const outgoing = underlay.value?.firstElementChild;
			const outgoingBase = outgoing?.style.transform && outgoing.style.transform !== "none" ? `${outgoing.style.transform} ` : "";
			const incomingTransform = (scale) => `${baseTransform}scale(${scale})`.trim();
			const outgoingTransformAt = (scale) => `${outgoingBase}scale(${scale})`.trim();
			underlay.value?.classList.add("is-flythrough");
			frames = transition.hasBounce ? [
				{
					transform: incomingTransform(direction === "in" ? 2.4 : .35),
					opacity: 0
				},
				{
					transform: incomingTransform(1.08),
					opacity: 1,
					offset: .82
				},
				{
					transform: incomingTransform(1),
					opacity: 1
				}
			] : [{
				transform: incomingTransform(direction === "in" ? 2.4 : .35),
				opacity: 0
			}, {
				transform: incomingTransform(1),
				opacity: 1
			}];
			outgoingFrames = direction === "in" ? [{
				transform: outgoingTransformAt(1),
				opacity: 1
			}, {
				transform: outgoingTransformAt(.35),
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
			const outgoing = underlay.value?.firstElementChild;
			const outgoingBase = outgoing?.style.transform && outgoing.style.transform !== "none" ? `${outgoing.style.transform} ` : "";
			const warped = (base, x, y, scale) => `${base}perspective(1200px) rotateX(${x}deg) rotateY(${y}deg) skewX(${y / 3}deg) scale(${scale})`.trim();
			if (direction === "out") {
				underlay.value?.classList.add("is-warping-out");
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
					transform: warped(outgoingBase, 18, -24, .72),
					opacity: 0,
					offset: 1
				}];
			} else frames = [
				{
					transform: warped(incomingBase, -18, 24, .72),
					opacity: 0
				},
				{
					transform: warped(incomingBase, 5, -7, 1.04),
					opacity: 1,
					offset: .78
				},
				{
					transform: incomingBase.trim() || "none",
					opacity: 1
				}
			];
		} else if (effect === "vortex") {
			const outgoing = underlay.value?.firstElementChild;
			const outgoingBase = outgoing?.style.transform && outgoing.style.transform !== "none" ? `${outgoing.style.transform} ` : "";
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
			underlay.value?.classList.add("is-vortexing");
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
		const outgoing = underlay.value?.firstElementChild;
		let outgoingAnimation;
		if (outgoingFrames && outgoing) outgoingAnimation = outgoing.animate(outgoingFrames, timing);
		if (effect === "fade" && transition.throughBlack) {
			shell.value?.classList.add("is-fading-through-black");
			outgoing?.animate([
				{
					opacity: 1,
					offset: 0
				},
				{
					opacity: 0,
					offset: .5
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
		if (revealThroughBlack) shell.value?.classList.add("is-fading-through-black");
		let animation;
		if (shredOutgoing && outgoing) animation = outgoing.animate(frames, timing);
		else if (shredOutgoing) {
			underlay.value?.classList.remove("is-shredding-out");
			underlay.value?.replaceChildren();
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
			underlay.value?.replaceChildren();
			return;
		} else if (effect === "split" && transition.direction === "in" && outgoing) {
			const center = transition.orientation === "horz" ? "inset(50% 0)" : "inset(0 50%)";
			animation = outgoing.animate([{ clipPath: "inset(0)" }, { clipPath: center }], timing);
		} else if (effect === "split" && transition.direction === "in") {
			underlay.value?.replaceChildren();
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
						offset: .5
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
			animation.cancel();
			outgoingAnimation?.cancel();
			if (slide.value?.id === activeSlide.id) {
				underlay.value?.replaceChildren();
				underlay.value?.classList.remove("is-shredding-out", "is-flythrough", "is-warping-out", "is-vortexing");
				shell.value?.classList.remove("is-fading-through-black");
			}
		};
	}, { flush: "post" });
}

export { SlideStage_default as S, cssFontFamily as c, paragraphEntries as p, scriptSegments as s, useSlideTransition as u };
//# sourceMappingURL=useSlideTransition-DcddA0qZ.mjs.map
