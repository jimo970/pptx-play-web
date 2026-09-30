function gridTransitionPathData(width: number, height: number, orientation: 'horz' | 'vert', progress: number, checker: boolean): string {
  // ponytail: fixed eight-cell short edge bounds path size; tune density from Office captures if mask fidelity needs it.
  const unit = Math.max(1, Math.min(width, height)) / 8
  const columns = Math.max(1, Math.round(width / unit))
  const rows = Math.max(1, Math.round(height / unit))
  const cellWidth = width / columns
  const cellHeight = height / rows
  const paths: string[] = []
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      let amount = progress
      if (checker) {
        const count = orientation === 'horz' ? columns : rows
        const position = orientation === 'horz' ? column : row
        const reversed = (orientation === 'horz' ? row : column) % 2 === 1
        const step = reversed ? count - 1 - position : position
        const revealWindow = 0.28
        amount = Math.max(0, Math.min(1, (progress - step / Math.max(1, count - 1) * (1 - revealWindow)) / revealWindow))
      }
      const left = column * cellWidth - 0.5
      const right = (column + 1) * cellWidth + 0.5
      const top = row * cellHeight - 0.5
      const bottom = (row + 1) * cellHeight + 0.5
      const halfWidth = checker || orientation === 'vert' ? cellWidth * amount / 2 : cellWidth / 2
      const halfHeight = checker || orientation === 'horz' ? cellHeight * amount / 2 : cellHeight / 2
      const centerX = (left + right) / 2
      const centerY = (top + bottom) / 2
      paths.push(`M${(centerX - halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} L${(centerX - halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} Z`)
    }
  }
  return paths.join(' ')
}

export function gridTransitionClipPath(width: number, height: number, orientation: 'horz' | 'vert', progress: number, checker: boolean): string {
  return `path("${gridTransitionPathData(width, height, orientation, progress, checker)}")`
}

export function gridTransitionMaskImage(width: number, height: number, orientation: 'horz' | 'vert', progress: number, checker: boolean): string {
  const maskWidth = Math.max(1, width)
  const maskHeight = Math.max(1, height)
  const path = gridTransitionPathData(maskWidth, maskHeight, orientation, progress, checker)
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${maskWidth}" height="${maskHeight}" viewBox="0 0 ${maskWidth} ${maskHeight}" preserveAspectRatio="none"><path fill="white" d="${path}"/></svg>`
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}

export function randomBarClipPath(width: number, height: number, orientation: 'horizontal' | 'vertical', progress: number, seed = 'slide'): string {
  // ponytail: 8 bars per short-edge unit bounds mask nodes; compare with Office captures before increasing density.
  const amount = Math.max(0, Math.min(progress, 1))
  const axisLength = orientation === 'horizontal' ? height : width
  const count = Math.max(1, Math.min(64, Math.ceil(axisLength / Math.max(1, Math.min(width, height) / 8))))
  const bandSize = axisLength / count
  let state = 2166136261
  for (const character of `${seed}:${orientation}`) state = Math.imul(state ^ character.charCodeAt(0), 16777619)
  const order = Array.from({ length: count }, (_, index) => index)
  for (let index = count - 1; index > 0; index--) {
    state ^= state << 13
    state ^= state >>> 17
    state ^= state << 5
    const swap = (state >>> 0) % (index + 1)
    const current = order[index]!
    order[index] = order[swap]!
    order[swap] = current
  }
  const window = 0.3
  const paths = order.map((rank, band) => {
    const start = rank / Math.max(1, count - 1) * (1 - window)
    const reveal = Math.max(0, Math.min(1, (amount - start) / window))
    const center = (band + 0.5) * bandSize
    const half = (bandSize * reveal + (reveal > 0 ? 1 : 0)) / 2
    if (orientation === 'horizontal') {
      return `M0 ${(center - half).toFixed(2)} H${width} V${(center + half).toFixed(2)} H0 Z`
    }
    return `M${(center - half).toFixed(2)} 0 H${(center + half).toFixed(2)} V${height} H${(center - half).toFixed(2)} Z`
  })
  return `path("${paths.join(' ')}")`
}

export function clipPathMaskImage(clipPath: string, width: number, height: number): string | undefined {
  const match = /^path\((?:(evenodd),\s*)?"([\s\S]*)"\)$/.exec(clipPath)
  if (!match) return undefined
  const maskWidth = Math.max(1, width)
  const maskHeight = Math.max(1, height)
  const fillRule = match[1] || 'nonzero'
  const path = match[2] || ''
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${maskWidth}" height="${maskHeight}" viewBox="0 0 ${maskWidth} ${maskHeight}" preserveAspectRatio="none"><path fill="white" fill-rule="${fillRule}" d="${path}"/></svg>`
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}

export function combClipPath(width: number, height: number, orientation: 'horz' | 'vert', progress: number): string {
  const amount = Math.max(0, Math.min(progress, 1))
  const axisLength = orientation === 'horz' ? width : height
  const crossAxisLength = orientation === 'horz' ? height : width
  const barCount = Math.max(1, Math.min(64, Math.ceil(crossAxisLength / Math.max(1, Math.min(width, height) / 8))))
  const barSize = crossAxisLength / barCount
  const paths = Array.from({ length: barCount }, (_, index) => {
    const start = index % 2 === 0 ? 0 : axisLength
    const end = index % 2 === 0 ? axisLength * amount : axisLength * (1 - amount)
    const crossStart = index * barSize - 0.5
    const crossEnd = (index + 1) * barSize + 0.5
    if (orientation === 'horz') return `M${Math.min(start, end).toFixed(2)} ${crossStart.toFixed(2)} H${Math.max(start, end).toFixed(2)} V${crossEnd.toFixed(2)} H${Math.min(start, end).toFixed(2)} Z`
    return `M${crossStart.toFixed(2)} ${Math.min(start, end).toFixed(2)} H${crossEnd.toFixed(2)} V${Math.max(start, end).toFixed(2)} H${crossStart.toFixed(2)} Z`
  })
  return `path("${paths.join(' ')}")`
}

export function stripsClipPath(width: number, height: number, direction: string, progress: number): string {
  const amount = Math.max(0, Math.min(progress, 1))
  const count = 8
  const bandHeight = height / count
  const reverseRows = direction.endsWith('u')
  const reverseColumns = direction.startsWith('l')
  const staggerWindow = 0.38
  const paths = Array.from({ length: count }, (_, band) => {
    const rank = reverseRows ? count - 1 - band : band
    const start = rank / (count - 1) * (1 - staggerWindow)
    const reveal = Math.max(0, Math.min(1, (amount - start) / staggerWindow))
    const top = band * bandHeight - 0.5
    const bottom = (band + 1) * bandHeight + 0.5
    const base = reverseColumns ? width : 0
    const travel = width * reveal
    const slant = reveal > 0 && reveal < 1 ? bandHeight * 0.75 : 0
    const first = reverseColumns ? base - travel : base + travel
    const leadingTop = Math.max(0, Math.min(width, first + (reverseColumns ? -slant : slant)))
    const leadingBottom = Math.max(0, Math.min(width, first - (reverseColumns ? -slant : slant)))
    const points: [number, number][] = reverseColumns
      ? [[leadingTop, top], [width, top], [width, bottom], [leadingBottom, bottom], [leadingTop, top]]
      : [[0, top], [leadingTop, top], [leadingBottom, bottom], [0, bottom], [0, top]]
    return `M${points.map(([x, y]) => `${x.toFixed(2)} ${y.toFixed(2)}`).join(' L')} Z`
  })
  return `path("${paths.join(' ')}")`
}

export function barnClipPath(orientation: 'horizontal' | 'vertical', motion: 'in' | 'out', progress: number): string {
  const amount = Math.max(0, Math.min(progress, 1))
  const inset = (motion === 'in' ? 1 - amount : amount) * 50
  return orientation === 'horizontal' ? `inset(${inset}% 0)` : `inset(0 ${inset}%)`
}

export function shapeEffectClipPath(shape: 'circle' | 'diamond' | 'box' | 'plus', width: number, height: number, progress: number, direction: 'in' | 'out'): string {
  // CSS path() with evenodd creates the exterior mask for Office's shape-out filters.
  // https://www.w3.org/TR/css-shapes/#funcdef-basic-shape-path
  const visible = Math.max(0, Math.min(progress, 1))
  const cx = width / 2, cy = height / 2
  const outer = `M0 0 H${width} V${height} H0 Z`
  if (shape === 'box') {
    if (direction === 'in') return `inset(${(1 - visible) * 50}% ${(1 - visible) * 50}%)`
    if (visible === 0) return 'inset(50%)'
    if (visible === 1) return 'inset(0)'
    const insetX = (1 - visible) * width / 2
    const insetY = (1 - visible) * height / 2
    return `path(evenodd, "${outer} M${insetX} ${insetY} H${width - insetX} V${height - insetY} H${insetX} Z")`
  }
  if (shape === 'plus') {
    if (direction === 'in') return visible === 0 ? 'inset(50%)' : visible === 1 ? 'inset(0)' : `path("${plusPath(width, height, visible)}")`
    if (visible === 1) return 'inset(0)'
    if (visible === 0) return 'inset(50%)'
    return `path(evenodd, "${outer} ${plusPath(width, height, 1 - visible)}")`
  }
  if (shape === 'diamond') {
    if (direction === 'in' && visible === 1) return 'inset(0)'
    const radiusX = width * visible, radiusY = height * visible
    if (visible === 0) return 'inset(50%)'
    const inner = `M${cx} ${cy - radiusY} L${cx + radiusX} ${cy} L${cx} ${cy + radiusY} L${cx - radiusX} ${cy} Z`
    return direction === 'in' ? `path("${inner}")` : visible === 1 ? 'inset(0)' : `path(evenodd, "${outer} ${inner}")`
  }
  const radius = Math.hypot(width, height) / 2 * (direction === 'in' ? visible : 1 - visible)
  if (radius <= 0) return direction === 'in' ? 'inset(50%)' : 'inset(0)'
  const circle = `M${cx} ${cy - radius} A${radius} ${radius} 0 1 1 ${cx} ${cy + radius} A${radius} ${radius} 0 1 1 ${cx} ${cy - radius} Z`
  return direction === 'in' ? `path("${circle}")` : radius >= Math.hypot(width, height) / 2 ? 'inset(50%)' : `path(evenodd, "${outer} ${circle}")`
}

export function rippleTransitionClipPath(width: number, height: number, direction: 'center' | 'lu' | 'ru' | 'ld' | 'rd', progress: number): string {
  // ponytail: 64 vertices keep the animated path small; use SVG displacement if Office captures show faceting.
  const amount = Math.max(0, Math.min(1, progress))
  const origins = {
    center: [width / 2, height / 2],
    lu: [width, height],
    ru: [0, height],
    ld: [width, 0],
    rd: [0, 0],
  } as const
  const [originX, originY] = origins[direction]
  const maxRadius = Math.max(...[
    [0, 0], [width, 0], [0, height], [width, height],
  ].map(([x, y]) => Math.hypot(x! - originX, y! - originY)))
  const amplitude = Math.min(width, height) * 0.025 * Math.sin(Math.PI * amount)
  const points = Array.from({ length: 64 }, (_, index) => {
    const angle = Math.PI * 2 * index / 64
    const radius = Math.max(0, maxRadius * amount + Math.sin(angle * 6 + amount * Math.PI * 4) * amplitude)
    return `${(originX + Math.cos(angle) * radius).toFixed(2)} ${(originY + Math.sin(angle) * radius).toFixed(2)}`
  })
  return `path("M${originX.toFixed(2)} ${originY.toFixed(2)} L${points.join(' L')} Z")`
}

export function glitterTransitionClipPath(width: number, height: number, direction: string, pattern: 'diamond' | 'hexagon', progress: number): string {
  // ponytail: about 220 tiles cap path size; tune density only from Office screenshot captures.
  if (width <= 0 || height <= 0) return 'inset(50%)'
  const amount = Math.max(0, Math.min(1, progress))
  const moveDirections: Record<string, readonly [number, number]> = {
    l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1],
    ld: [-Math.SQRT1_2, Math.SQRT1_2], lu: [-Math.SQRT1_2, -Math.SQRT1_2],
    rd: [Math.SQRT1_2, Math.SQRT1_2], ru: [Math.SQRT1_2, -Math.SQRT1_2],
  }
  const [moveX, moveY] = moveDirections[direction] || moveDirections.l!
  const startX = -moveX, startY = -moveY
  const corners = [[0, 0], [width, 0], [0, height], [width, height]]
  const projections = corners.map(([x, y]) => x! * startX + y! * startY)
  const minProjection = Math.min(...projections)
  const projectionSpan = Math.max(...projections) - minProjection || 1
  const paths: string[] = []
  const revealTile = (cx: number, cy: number, vertices: readonly (readonly [number, number])[]) => {
    const projection = cx * startX + cy * startY
    const rank = 1 - (projection - minProjection) / projectionSpan
    const local = Math.max(0, Math.min(1, (amount - rank * 0.72) / 0.28))
    paths.push(`M${cx.toFixed(2)} ${cy.toFixed(2)} ${vertices.map(([dx, dy]) => `L${(cx + dx * local).toFixed(2)} ${(cy + dy * local).toFixed(2)}`).join(' ')} Z`)
  }

  if (pattern === 'diamond') {
    const radius = Math.sqrt(width * height / 440)
    const rows = Math.ceil(height / radius) + 3
    const columns = Math.ceil(width / radius) + 3
    for (let row = -rows; row <= rows; row++) {
      for (let column = -columns; column <= columns; column++) {
        if ((row + column) % 2 !== 0) continue
        const cx = width / 2 + column * radius
        const cy = height / 2 + row * radius
        if (cx + radius < 0 || cx - radius > width || cy + radius < 0 || cy - radius > height) continue
        revealTile(cx, cy, [[0, -radius], [radius, 0], [0, radius], [-radius, 0]])
      }
    }
  } else {
    const side = Math.sqrt(width * height / (220 * 2.6))
    const halfWidth = Math.sqrt(3) * side / 2
    const rowStep = side * 1.5
    const columnStep = halfWidth * 2
    const vertices: readonly (readonly [number, number])[] = [[0, -side], [halfWidth, -side / 2], [halfWidth, side / 2], [0, side], [-halfWidth, side / 2], [-halfWidth, -side / 2]]
    for (let row = -Math.ceil(height / rowStep) - 2; row <= Math.ceil(height / rowStep) + 2; row++) {
      for (let column = -Math.ceil(width / columnStep) - 2; column <= Math.ceil(width / columnStep) + 2; column++) {
        const cx = width / 2 + columnStep * (column + row / 2)
        const cy = height / 2 + rowStep * row
        if (cx + halfWidth < 0 || cx - halfWidth > width || cy + side < 0 || cy - side > height) continue
        revealTile(cx, cy, vertices)
      }
    }
  }
  return `path("${paths.join(' ')}")`
}

export function shredTransitionClipPath(width: number, height: number, pattern: 'strip' | 'rectangle', progress: number): string {
  if (width <= 0 || height <= 0) return 'inset(50%)'
  const amount = Math.max(0, Math.min(1, progress))
  const columns = pattern === 'strip' ? 16 : 10
  const rows = pattern === 'strip' ? 1 : 7
  const cellWidth = width / columns
  const cellHeight = height / rows
  const paths: string[] = []

  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      const rank = pattern === 'strip'
        ? column / Math.max(1, columns - 1)
        : (row + column) / (rows + columns - 2)
      const local = Math.max(0, Math.min(1, (amount - rank * 0.68) / 0.32))
      const overlap = local ? 0.25 : 0
      const halfWidth = cellWidth * local / 2 + overlap
      const halfHeight = cellHeight * local / 2 + overlap
      const centerX = (column + 0.5) * cellWidth
      const centerY = (row + 0.5) * cellHeight
      paths.push(`M${(centerX - halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY - halfHeight).toFixed(2)} L${(centerX + halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} L${(centerX - halfWidth).toFixed(2)} ${(centerY + halfHeight).toFixed(2)} Z`)
    }
  }

  return `path("${paths.join(' ')}")`
}

function plusPath(width: number, height: number, amount: number): string {
  const cx = width / 2, cy = height / 2
  const reachX = width * amount / 2
  const reachY = height * amount / 2
  const arm = amount * 0.28 + Math.pow(amount, 12) * 0.72
  const armX = width * arm / 2
  const armY = height * arm / 2
  return `M${cx - armX} ${cy - reachY} H${cx + armX} V${cy - armY} H${cx + reachX} V${cy + armY} H${cx + armX} V${cy + reachY} H${cx - armX} V${cy + armY} H${cx - reachX} V${cy - armY} H${cx - armX} Z`
}

export function plusTransitionClipPath(width: number, height: number, progress: number): string {
  return `path("${plusPath(width, height, Math.max(0, Math.min(1, progress)))}")`
}

export function wedgeTransitionClipPath(width: number, height: number, progress: number): string {
  const amount = Math.max(0, Math.min(1, progress))
  const centerX = width / 2
  const centerY = height / 2
  const radius = Math.hypot(width, height) / 2 + 1
  const halfAngle = Math.PI * amount
  const points = Array.from({ length: 13 }, (_, point) => {
    const angle = -Math.PI / 2 - halfAngle + halfAngle * 2 * point / 12
    return `L${(centerX + Math.cos(angle) * radius).toFixed(2)} ${(centerY + Math.sin(angle) * radius).toFixed(2)}`
  })
  return `path("M${centerX.toFixed(2)} ${centerY.toFixed(2)} ${points.join(' ')} Z")`
}

export function dissolveClipPath(width: number, height: number, progress: number, seed = 'slide'): string {
  // ponytail: 8 tiles on the short edge bounds path size; replace with a pixel mask if frame captures show the coarse dissolve is noticeable.
  const amount = Math.max(0, Math.min(progress, 1))
  if (amount === 0) return 'inset(50%)'
  if (amount === 1) return 'inset(0)'
  const unit = Math.max(1, Math.min(width, height)) / 8
  const columns = Math.max(1, Math.round(width / unit))
  const rows = Math.max(1, Math.round(height / unit))
  const cellWidth = width / columns
  const cellHeight = height / rows
  let seedHash = 2166136261
  for (const character of seed) seedHash = Math.imul(seedHash ^ character.charCodeAt(0), 16777619)
  const paths: string[] = []
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      let hash = Math.imul(row + 1, 374761393) ^ Math.imul(column + 1, 668265263) ^ seedHash
      hash = Math.imul(hash ^ (hash >>> 13), 1274126177)
      const visible = ((hash ^ (hash >>> 16)) >>> 0) / 0x1_0000_0000 < amount
      const left = column * cellWidth - 0.5
      const right = (column + 1) * cellWidth + 0.5
      const top = row * cellHeight - 0.5
      const bottom = (row + 1) * cellHeight + 0.5
      const cx = (left + right) / 2
      const cy = (top + bottom) / 2
      paths.push(visible
        ? `M${left} ${top} L${right} ${top} L${right} ${bottom} L${left} ${bottom} Z`
        : `M${cx} ${cy} L${cx} ${cy} L${cx} ${cy} L${cx} ${cy} Z`)
    }
  }
  return `path("${paths.join(' ')}")`
}

export function wheelClipPath(width: number, height: number, spokes: number, progress: number, reverse = false): string {
  const centerX = width / 2
  const centerY = height / 2
  const radius = Math.hypot(width, height) / 2 + 1
  const pointsPerSpoke = 12
  const paths = Array.from({ length: spokes }, (_, spoke) => {
    const centerAngle = -Math.PI / 2 + (spoke + 0.5) * Math.PI * 2 / spokes
    const halfAngle = Math.PI / spokes * progress
    const points = Array.from({ length: pointsPerSpoke + 1 }, (_, point) => {
      const angle = reverse
        ? -Math.PI / 2 + (spoke + 1) * Math.PI * 2 / spokes - Math.PI * 2 / spokes * progress * point / pointsPerSpoke
        : centerAngle - halfAngle + halfAngle * 2 * point / pointsPerSpoke
      return `L${(centerX + Math.cos(angle) * radius).toFixed(2)} ${(centerY + Math.sin(angle) * radius).toFixed(2)}`
    })
    return `M${centerX.toFixed(2)} ${centerY.toFixed(2)} ${points.join(' ')} Z`
  })
  return `path("${paths.join(' ')}")`
}
