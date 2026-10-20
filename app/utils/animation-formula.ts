const ANIMATION_FORMULA_FUNCTIONS: Record<string, (...args: number[]) => number> = {
    abs: Math.abs,
    acos: Math.acos,
    asin: Math.asin,
    atan: Math.atan,
    ceil: Math.ceil,
    cos: Math.cos,
    cosh: Math.cosh,
    deg: value => (value * 180) / Math.PI,
    exp: Math.exp,
    floor: Math.floor,
    ln: Math.log,
    max: Math.max,
    min: Math.min,
    rad: value => (value * Math.PI) / 180,
    rand: maximum => maximum < 0 ? Number.NaN : Math.random() * maximum,
    sin: Math.sin,
    sinh: Math.sinh,
    sqrt: Math.sqrt,
    tan: Math.tan,
    tanh: Math.tanh
}

function normalizedFormulaVariables(variables: Record<string, number>): Map<string, number> {
    return new Map(Object.entries(variables).map(([name, value]) => [name.toLowerCase(), value]))
}

export function evaluateAnimationFormula(formula: string, variables: Record<string, number>, allowBareVariables = false): number | undefined {
    return evaluateAnimationFormulaWithVariables(formula, variables, normalizedFormulaVariables(variables), allowBareVariables)
}

function evaluateAnimationFormulaWithVariables(
    formula: string,
    variables: Record<string, number>,
    normalizedVariables: Map<string, number>,
    allowBareVariables = false,
    progressValue?: number,
    randomValue: () => number = Math.random
): number | undefined {
    if (!formula.trim() || formula.length > 16_384) return undefined
    const tokens: string[] = []
    const tokenizer = /#[A-Za-z][A-Za-z0-9_.]*|\$|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[A-Za-z][A-Za-z0-9_.]*|[()+\-*/%^,]/gy
    let scan = 0
    while (scan < formula.length) {
        if (/\s/.test(formula[scan] || '')) {
            scan++
            continue
        }
        tokenizer.lastIndex = scan
        const match = tokenizer.exec(formula)
        if (!match) return undefined
        tokens.push(match[0])
        if (tokens.length > 512) return undefined
        scan = tokenizer.lastIndex
    }

    // Formula ATTRIBUTE references use attrName values, which PowerPoint treats as case-insensitive.
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/baf99b7b-b2a7-44bb-a501-f6147d6c7123
    const variableValue = (name: string) => Object.hasOwn(variables, name) ? variables[name] : normalizedVariables.get(name.toLowerCase())
    let cursor = 0
    const finite = (value: number | undefined): value is number => value !== undefined && Number.isFinite(value)
    const formulaValue = (): number | undefined => {
        let value = term()
        while (value !== undefined && (tokens[cursor] === '+' || tokens[cursor] === '-')) {
            const operator = tokens[cursor++]
            const right = term()
            if (right === undefined) return undefined
            value = operator === '+' ? value + right : value - right
            if (!Number.isFinite(value)) return undefined
        }
        return value
    }
    const factor = (depth: number): number | undefined => {
        if (depth > 255) return undefined
        const token = tokens[cursor++]
        if (token === undefined) return undefined
        if (token === '+' || token === '-') {
            const value = factor(depth + 1)
            return value === undefined ? undefined : token === '-' ? -value : value
        }
        if (token === '(') {
            const value = formulaValue()
            if (!finite(value) || tokens[cursor++] !== ')') return undefined
            return value
        }
        if (token.startsWith('#')) return variableValue(token.slice(1))
        if (token === '$') return progressValue ?? variables.$
        if (token === 'pi') return Math.PI
        if (token === 'e') return Math.E
        const numeric = Number(token)
        if (Number.isFinite(numeric)) return numeric
        if (allowBareVariables) {
            const variable = variableValue(token)
            if (variable !== undefined) return variable
        }
        const fn = token === 'rand'
            ? (maximum: number) => maximum < 0 ? Number.NaN : randomValue() * maximum
            : Object.hasOwn(ANIMATION_FORMULA_FUNCTIONS, token) ? ANIMATION_FORMULA_FUNCTIONS[token] : undefined
        if (!fn || tokens[cursor++] !== '(') return undefined
        const first = formulaValue()
        if (!finite(first)) return undefined
        let args = [first]
        if (token === 'max' || token === 'min') {
            if (tokens[cursor++] !== ',') return undefined
            const second = formulaValue()
            if (!finite(second)) return undefined
            args.push(second)
        }
        if (tokens[cursor++] !== ')') return undefined
        const result = fn(...args)
        return Number.isFinite(result) ? result : undefined
    }
    const power = (): number | undefined => {
        let value = factor(0)
        while (value !== undefined && tokens[cursor] === '^') {
            cursor++
            const right = factor(0)
            if (right === undefined) return undefined
            value = value ** right
            if (!Number.isFinite(value)) return undefined
        }
        return value
    }
    const term = (): number | undefined => {
        let value = power()
        while (value !== undefined && (tokens[cursor] === '*' || tokens[cursor] === '/' || tokens[cursor] === '%')) {
            const operator = tokens[cursor++]
            const right = power()
            if (right === undefined) return undefined
            value = operator === '*' ? value * right : operator === '/' ? value / right : value % right
            if (!Number.isFinite(value)) return undefined
        }
        return value
    }

    const value = formulaValue()
    return finite(value) && cursor === tokens.length ? value : undefined
}

export interface AnimationFormulaFrame {
    offset: number
    value: number
    formula?: string
}

export function sampleAnimationFormula(
    keyframes: AnimationFormulaFrame[],
    durationMs: number,
    variables: Record<string, number>,
    minimum: number,
    maximum: number,
    keyframeMode: 'lin' | 'discrete' = 'lin',
    randomSeed?: number
): Array<{offset: number; value: number}> | undefined {
    // Formula frames use normalized offsets, so an indefinite animation can still be sampled
    // at the existing 900-frame cap while playback itself remains at its initial progress.
    if (!keyframes.length || keyframes.length > 256 || Number.isNaN(durationMs) || durationMs < 0 ||
        !Number.isFinite(minimum) || !Number.isFinite(maximum) || minimum > maximum ||
        keyframes.some((frame, index) => !Number.isFinite(frame.value) || !Number.isFinite(frame.offset) ||
            frame.offset < 0 || frame.offset > 1 || index > 0 && frame.offset <= keyframes[index - 1]!.offset)) return undefined
    const sampleCount = Math.min(900, Math.max(60, Math.ceil(durationMs / 1000 * 60)))
    const samples: Array<{offset: number; value: number}> = []
    const normalizedVariables = normalizedFormulaVariables(variables)
    let randomState = randomSeed === undefined ? undefined : randomSeed >>> 0
    const randomValue = () => {
        if (randomState === undefined) return Math.random()
        randomState = (Math.imul(randomState, 1_664_525) + 1_013_904_223) >>> 0
        return randomState / 0x1_0000_0000
    }
    let rightIndex = 0
    let formulaIndex = 0
    let activeFormula: string | undefined
    for (let index = 0; index <= sampleCount; index++) {
        const progress = index / sampleCount
        while (rightIndex < keyframes.length && keyframes[rightIndex]!.offset < progress) rightIndex++
        while (formulaIndex < keyframes.length && keyframes[formulaIndex]!.offset <= progress) {
            const frame = keyframes[formulaIndex++]!
            if (frame.formula !== undefined) activeFormula = frame.formula
        }
        const frameAtProgress = keyframes[rightIndex]
        const isExactFrame = frameAtProgress?.offset === progress
        const left = rightIndex >= keyframes.length
            ? keyframes.at(-1)!
            : keyframes[Math.max(0, rightIndex - (isExactFrame ? 0 : 1))]!
        const right = frameAtProgress || left
        const ratio = right.offset === left.offset ? 0 : (progress - left.offset) / (right.offset - left.offset)
        const input = keyframeMode === 'discrete' ? left.value : left.value + (right.value - left.value) * Math.max(0, Math.min(1, ratio))
        // Formulas consume the interpolated tav value; geometric variables remain static.
        // MS-OE376 tav's Bounce example uses input 0.5 -> 1; LibreOffice also applies
        // its formula after value interpolation in ValuesActivity::perform.
        const value = activeFormula ? evaluateAnimationFormulaWithVariables(activeFormula, variables, normalizedVariables, false, input, randomValue) : input
        if (value === undefined || value < minimum || value > maximum) return undefined
        samples.push({offset: progress, value})
    }
    return samples
}

export function animationGeometryVariables(
    element: {x: number; y: number; width: number; height: number; renderMatrix?: [number, number, number, number, number, number]},
    slideWidth: number,
    slideHeight: number
): Record<'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h', number> {
    const matrix = element.renderMatrix
    return {
        ppt_x: (element.x + (matrix ? matrix[0] * element.width / 2 + matrix[2] * element.height / 2 + matrix[4] : element.width / 2)) / slideWidth,
        ppt_y: (element.y + (matrix ? matrix[1] * element.width / 2 + matrix[3] * element.height / 2 + matrix[5] : element.height / 2)) / slideHeight,
        ppt_w: element.width * (matrix ? Math.hypot(matrix[0], matrix[1]) : 1) / slideWidth,
        ppt_h: element.height * (matrix ? Math.hypot(matrix[2], matrix[3]) : 1) / slideHeight
    }
}
