// Univer <-> Excel 通用转换工具：坐标、颜色、样式、尺寸

// ---------- 坐标 ----------
export function colToLetters(col) {
  let s = ''
  let n = col + 1
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function lettersToCol(letters) {
  let col = 0
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return col - 1
}

export const cellRef = (row, col, abs = false) =>
  abs ? `$${colToLetters(col)}$${row + 1}` : `${colToLetters(col)}${row + 1}`

export function rangeToA1(range, abs = false) {
  const a = cellRef(range.startRow, range.startColumn, abs)
  const b = cellRef(range.endRow, range.endColumn, abs)
  return a === b ? a : `${a}:${b}`
}

// 'A1' / 'A1:B2' / '$A$1:$B$2'（整行整列 'A:A' / '1:1' 也支持）
export function a1ToRange(ref) {
  const clean = ref.replace(/\$/g, '').trim()
  const [a, b = a] = clean.split(':')
  const parse = (s, isEnd) => {
    const m = s.match(/^([A-Z]*)(\d*)$/i)
    if (!m) throw new Error(`无法解析区域 ${ref}`)
    return {
      col: m[1] ? lettersToCol(m[1]) : (isEnd ? 16383 : 0),
      row: m[2] ? Number(m[2]) - 1 : (isEnd ? 1048575 : 0),
    }
  }
  const s = parse(a, false)
  const e = parse(b, true)
  return {
    startRow: Math.min(s.row, e.row),
    startColumn: Math.min(s.col, e.col),
    endRow: Math.max(s.row, e.row),
    endColumn: Math.max(s.col, e.col),
  }
}

// 'A1:B2 C3:D4' -> ranges
export const sqrefToRanges = (sqref) => String(sqref).trim().split(/\s+/).filter(Boolean).map(a1ToRange)
export const rangesToSqref = (ranges) => ranges.map((r) => rangeToA1(r)).join(' ')

export function clampRange(range, maxRow, maxCol) {
  return {
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: Math.min(range.endRow, maxRow),
    endColumn: Math.min(range.endColumn, maxCol),
  }
}

// 把一组单元格坐标合并成尽量少的矩形（先按列连续行合并，再合并相同行段的相邻列）
export function cellsToRanges(cells) {
  const byCol = new Map()
  for (const { row, col } of cells) {
    if (!byCol.has(col)) byCol.set(col, [])
    byCol.get(col).push(row)
  }
  const segments = []
  for (const [col, rows] of byCol) {
    rows.sort((a, b) => a - b)
    let start = rows[0]
    let prev = rows[0]
    for (let i = 1; i <= rows.length; i++) {
      if (i < rows.length && rows[i] === prev + 1) { prev = rows[i]; continue }
      segments.push({ startRow: start, endRow: prev, startColumn: col, endColumn: col })
      start = rows[i]
      prev = rows[i]
    }
  }
  segments.sort((a, b) => a.startRow - b.startRow || a.endRow - b.endRow || a.startColumn - b.startColumn)
  const out = []
  for (const s of segments) {
    const last = out[out.length - 1]
    if (last && last.startRow === s.startRow && last.endRow === s.endRow && last.endColumn + 1 === s.startColumn) last.endColumn = s.endColumn
    else out.push({ ...s })
  }
  return out
}

// 求一组矩形在 [0..maxRow]×[0..maxCol] 范围内的补集，返回尽量少的矩形
export function complementRects(rects, maxRow, maxCol) {
  const clip = rects
    .map((r) => ({ startRow: Math.max(0, r.startRow), endRow: Math.min(maxRow, r.endRow), startColumn: Math.max(0, r.startColumn), endColumn: Math.min(maxCol, r.endColumn) }))
    .filter((r) => r.startRow <= r.endRow && r.startColumn <= r.endColumn)
  // 按行切成横条：条内每一行被覆盖的列区间相同
  const cuts = new Set([0, maxRow + 1])
  for (const r of clip) { cuts.add(r.startRow); cuts.add(r.endRow + 1) }
  const rows = [...cuts].sort((a, b) => a - b)
  const bands = []
  for (let i = 0; i < rows.length - 1; i++) {
    const top = rows[i]
    const bottom = rows[i + 1] - 1
    if (top > bottom) continue
    const covered = clip.filter((r) => r.startRow <= top && r.endRow >= top).map((r) => [r.startColumn, r.endColumn]).sort((a, b) => a[0] - b[0])
    const gaps = []
    let c = 0
    for (const [s, e] of covered) {
      if (s > c) gaps.push([c, s - 1])
      c = Math.max(c, e + 1)
    }
    if (c <= maxCol) gaps.push([c, maxCol])
    bands.push({ top, bottom, gaps })
  }
  // 相邻横条里列区间相同的合并成一个矩形
  const out = []
  const open = new Map()
  for (const band of bands) {
    const keys = new Set(band.gaps.map((g) => g.join(',')))
    for (const [key, rect] of open) {
      if (!keys.has(key) || rect.endRow !== band.top - 1) { out.push(rect); open.delete(key) }
    }
    for (const [s, e] of band.gaps) {
      const key = `${s},${e}`
      if (open.has(key)) open.get(key).endRow = band.bottom
      else open.set(key, { startRow: band.top, endRow: band.bottom, startColumn: s, endColumn: e })
    }
  }
  out.push(...open.values())
  return out
}

// ---------- 颜色 ----------
// Univer 的颜色可能是 '#RRGGBB' / '#RGB' / 'rgb(r,g,b)' / 'rgba(...)'，统一转 Excel ARGB
export function toArgb(input) {
  const rgb = typeof input === 'string' ? input : input?.rgb
  if (!rgb) return undefined
  const s = String(rgb).trim()
  let m = s.match(/^#?([0-9a-f]{6})$/i)
  if (m) return 'FF' + m[1].toUpperCase()
  m = s.match(/^#?([0-9a-f]{8})$/i)
  if (m) return 'FF' + m[1].slice(0, 6).toUpperCase() // #RRGGBBAA：Excel 单元格颜色不支持透明度，忽略 AA
  m = s.match(/^#?([0-9a-f]{3})$/i)
  if (m) return 'FF' + m[1].split('').map((c) => c + c).join('').toUpperCase()
  m = s.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i)
  if (m) return 'FF' + m.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase()
  return undefined
}

// Excel 默认主题色（文件里没有主题时兜底），索引顺序：lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

// Excel 旧式索引色 0-63
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

export function parseThemeColors(workbook) {
  const xml = workbook._themes?.theme1
  if (!xml) return DEFAULT_THEME
  const pick = (tag) => {
    const block = xml.match(new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`))
    if (!block) return null
    const m = block[1].match(/(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"/)
    return m ? m[1].toUpperCase() : null
  }
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
  return order.map((tag, i) => pick(tag) || DEFAULT_THEME[i])
}

function applyTint(hex, tint) {
  if (!tint) return hex
  let r = parseInt(hex.slice(0, 2), 16) / 255
  let g = parseInt(hex.slice(2, 4), 16) / 255
  let b = parseInt(hex.slice(4, 6), 16) / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  let l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  if (s === 0) { r = g = b = l } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    r = hue2rgb(p, q, h + 1 / 3); g = hue2rgb(p, q, h); b = hue2rgb(p, q, h - 1 / 3)
  }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase()
}

// Excel 颜色对象 {argb} / {theme, tint} / {indexed} -> '#RRGGBB'
export function excelColorToHex(color, theme) {
  if (!color) return undefined
  let hex
  if (color.argb) hex = color.argb.length === 8 ? color.argb.slice(2) : color.argb
  else if (color.theme !== undefined) hex = theme[color.theme]
  else if (color.indexed !== undefined) hex = INDEXED[color.indexed]
  if (!hex) return undefined
  return '#' + applyTint(hex, color.tint)
}

// ---------- 单元格样式 ----------
const BORDER_TO_EXCEL = {
  1: 'thin', 2: 'hair', 3: 'dotted', 4: 'dashed', 5: 'dashDot', 6: 'dashDotDot',
  7: 'double', 8: 'medium', 9: 'mediumDashed', 10: 'mediumDashDot',
  11: 'mediumDashDotDot', 12: 'slantDashDot', 13: 'thick',
}
const BORDER_TO_UNIVER = Object.fromEntries(Object.entries(BORDER_TO_EXCEL).map(([k, v]) => [v, Number(k)]))
const H_TO_EXCEL = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify', 5: 'justify', 6: 'distributed' }
const H_TO_UNIVER = { left: 1, center: 2, right: 3, justify: 4, distributed: 6, centerContinuous: 2, fill: 1 }
const V_TO_EXCEL = { 1: 'top', 2: 'middle', 3: 'bottom' }
const V_TO_UNIVER = { top: 1, middle: 2, bottom: 3, justify: 2, distributed: 2 }

// Univer 字体样式 -> ExcelJS font（单元格样式、富文本片段、条件格式共用）
export function univerFontToExcel(style) {
  const font = {}
  if (!style) return font
  if (style.ff) font.name = style.ff
  if (style.fs) font.size = style.fs
  if (style.bl === 1) font.bold = true
  if (style.it === 1) font.italic = true
  if (style.ul?.s === 1) font.underline = true
  if (style.st?.s === 1) font.strike = true
  if (style.va === 2) font.vertAlign = 'subscript'
  if (style.va === 3) font.vertAlign = 'superscript'
  const fc = toArgb(style.cl)
  if (fc) font.color = { argb: fc }
  return font
}

export function excelFontToUniver(font, theme, { keepBlack = false } = {}) {
  const style = {}
  if (!font) return style
  if (font.name) style.ff = font.name
  if (font.size) style.fs = font.size
  if (font.bold) style.bl = 1
  if (font.italic) style.it = 1
  if (font.underline) style.ul = { s: 1 }
  if (font.strike) style.st = { s: 1 }
  if (font.vertAlign === 'subscript') style.va = 2
  if (font.vertAlign === 'superscript') style.va = 3
  const cl = excelColorToHex(font.color, theme)
  if (cl && (keepBlack || cl !== '#000000')) style.cl = { rgb: cl }
  return style
}

// Univer 单元格样式 -> ExcelJS cell 样式属性
export function applyUniverStyleToExcelCell(cell, style) {
  if (!style) return
  const font = univerFontToExcel(style)
  if (Object.keys(font).length) cell.font = font

  const bg = toArgb(style.bg)
  if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }

  const alignment = {}
  if (H_TO_EXCEL[style.ht]) alignment.horizontal = H_TO_EXCEL[style.ht]
  if (V_TO_EXCEL[style.vt]) alignment.vertical = V_TO_EXCEL[style.vt]
  if (style.tb === 3) alignment.wrapText = true
  if (style.tr?.v === 1) alignment.textRotation = 'vertical'
  else if (style.tr?.a) alignment.textRotation = style.tr.a
  if (style.pd?.l && style.pd.l >= 8) alignment.indent = Math.round(style.pd.l / 8)
  if (Object.keys(alignment).length) cell.alignment = alignment

  if (style.bd) {
    const border = {}
    for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) {
      const b = style.bd[k]
      if (b && BORDER_TO_EXCEL[b.s]) border[side] = { style: BORDER_TO_EXCEL[b.s], color: { argb: toArgb(b.cl) || 'FF000000' } }
    }
    if (Object.keys(border).length) cell.border = border
  }

  if (style.n?.pattern) cell.numFmt = style.n.pattern
}

// ExcelJS cell -> Univer 样式对象
export function excelCellStyleToUniver(cell, theme) {
  const style = excelFontToUniver(cell.font, theme)
  const fill = cell.fill
  if (fill?.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    const bg = excelColorToHex(fill.fgColor, theme) || excelColorToHex(fill.bgColor, theme)
    if (bg) style.bg = { rgb: bg }
  } else if (fill?.type === 'gradient' && fill.stops?.length) {
    const bg = excelColorToHex(fill.stops[0].color, theme)
    if (bg) style.bg = { rgb: bg }
  }
  const al = cell.alignment
  if (al) {
    if (H_TO_UNIVER[al.horizontal]) style.ht = H_TO_UNIVER[al.horizontal]
    if (V_TO_UNIVER[al.vertical]) style.vt = V_TO_UNIVER[al.vertical]
    if (al.wrapText) style.tb = 3
    if (al.textRotation === 'vertical') style.tr = { a: 0, v: 1 }
    else if (typeof al.textRotation === 'number' && al.textRotation) style.tr = { a: al.textRotation > 90 ? 90 - al.textRotation : al.textRotation, v: 0 }
    if (al.indent) style.pd = { l: al.indent * 8 }
  }
  const bd = cell.border
  if (bd) {
    const out = {}
    for (const [side, k] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']]) {
      const b = bd[side]
      if (b?.style && BORDER_TO_UNIVER[b.style]) out[k] = { s: BORDER_TO_UNIVER[b.style], cl: { rgb: excelColorToHex(b.color, theme) || '#000000' } }
    }
    if (Object.keys(out).length) style.bd = out
  }
  if (cell.numFmt && cell.numFmt !== 'General') style.n = { pattern: cell.numFmt }
  return Object.keys(style).length ? style : null
}

// 条件格式里的高亮样式（Excel dxf）<-> Univer IStyleBase
export function univerHighlightToDxf(style = {}) {
  const dxf = {}
  const font = univerFontToExcel(style)
  delete font.name
  delete font.size
  if (Object.keys(font).length) dxf.font = font
  const bg = toArgb(style.bg)
  if (bg) dxf.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg }, bgColor: { argb: bg } }
  return dxf
}

export function dxfToUniverHighlight(dxf, theme) {
  const style = excelFontToUniver(dxf?.font, theme, { keepBlack: true })
  delete style.ff
  delete style.fs
  const fill = dxf?.fill
  if (fill?.type === 'pattern') {
    const bg = excelColorToHex(fill.bgColor, theme) || excelColorToHex(fill.fgColor, theme)
    if (bg) style.bg = { rgb: bg }
  }
  return style
}

// ---------- 尺寸 ----------
// Univer 用像素；Excel 列宽用“字符数”（约 7px/字符），行高用磅（0.75pt/px）
export const pxToColWidth = (px) => Math.round((px / 7) * 100) / 100
export const colWidthToPx = (w) => Math.round(w * 7)
export const pxToPt = (px) => Math.round(px * 0.75 * 100) / 100
export const ptToPx = (pt) => Math.round(pt / 0.75)
export const EMU_PER_PX = 9525

// 按工作表的行高列宽，把 (行/列 + 偏移) 和像素坐标互相换算
export function createSheetMetrics(sheet) {
  const defaultW = sheet.defaultColumnWidth || 88
  const defaultH = sheet.defaultRowHeight || 24
  const colW = (c) => {
    const d = sheet.columnData?.[c]
    if (d?.hd === 1) return 0
    return d?.w ?? defaultW
  }
  const rowH = (r) => {
    const d = sheet.rowData?.[r]
    if (d?.hd === 1) return 0
    return d?.ah && d?.ia !== 0 ? d.ah : (d?.h ?? defaultH)
  }
  const colLeft = (c) => { let x = 0; for (let i = 0; i < c; i++) x += colW(i); return x }
  const rowTop = (r) => { let y = 0; for (let i = 0; i < r; i++) y += rowH(i); return y }
  // 像素 -> {col, colOffset}
  const locateX = (x) => { let c = 0; let left = 0; while (c < 16383 && left + colW(c) <= x) { left += colW(c); c++ } return { col: c, offset: x - left } }
  const locateY = (y) => { let r = 0; let top = 0; while (r < 1048575 && top + rowH(r) <= y) { top += rowH(r); r++ } return { row: r, offset: y - top } }
  return { colW, rowH, colLeft, rowTop, locateX, locateY }
}

// ---------- 其他 ----------
const EXCEL_EPOCH = Date.UTC(1899, 11, 30)
export const dateToSerial = (d) => (d.getTime() - EXCEL_EPOCH) / 86400000
export const serialToDate = (n) => new Date(EXCEL_EPOCH + n * 86400000)
export const formatDate = (d) => d.toISOString().slice(0, 10)

export function randomId(len = 12) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
  let s = ''
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)]
  return s
}

export function bytesToBase64(bytes) {
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk))
  return btoa(binary)
}

export function parseResources(snapshot) {
  const map = {}
  for (const r of snapshot.resources || []) {
    if (!r.data) continue
    try { map[r.name] = JSON.parse(r.data) } catch { /* ignore */ }
  }
  return map
}

// Univer 的单元格 p（文档）里的纯文本
export function docToText(p) {
  const stream = p?.body?.dataStream
  if (!stream) return ''
  return stream.replace(/\r\n$/, '').replace(/[\b]/g, '').replace(/\r/g, '\n')
}

// 生成一个最小的 Univer 单元格文档（富文本 / 超链接用）
export function createCellDoc(text, { textRuns = [], customRanges = [] } = {}) {
  const stream = text.replace(/\r?\n/g, '\r') + '\r\n'
  const paragraphs = []
  for (let i = 0; i < stream.length; i++) if (stream[i] === '\r' && stream[i + 1] !== '\n') paragraphs.push({ startIndex: i })
  paragraphs.push({ startIndex: stream.length - 2 })
  return {
    id: 'd',
    documentStyle: {},
    body: {
      dataStream: stream,
      textRuns,
      paragraphs,
      sectionBreaks: [{ startIndex: stream.length - 1 }],
      customRanges,
      customDecorations: [],
    },
  }
}
