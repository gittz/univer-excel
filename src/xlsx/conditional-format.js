// 条件格式：Univer 规则 <-> Excel 规则（ExcelJS conditionalFormattings 模型）
import {
  cellRef, rangeToA1, rangesToSqref, sqrefToRanges, toArgb, excelColorToHex,
  univerHighlightToDxf, dxfToUniverHighlight, randomId,
} from './utils.js'
import { toExcelFormula, fromExcelFormula } from './formula.js'

const NUM_OPS = new Set(['greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'equal', 'notEqual', 'between', 'notBetween'])
const OPPOSITE = { greaterThan: 'lessThanOrEqual', greaterThanOrEqual: 'lessThan', lessThan: 'greaterThanOrEqual', lessThanOrEqual: 'greaterThan' }

const ICON_SETS = new Set(['3Arrows', '3ArrowsGray', '4Arrows', '4ArrowsGray', '5Arrows', '5ArrowsGray', '3Triangles', '3TrafficLights1', '3Signs', '3TrafficLights2', '4RedToBlack', '4TrafficLights', '3Symbols', '3Symbols2', '3Flags', '4Rating', '5Rating', '5Quarters', '5Boxes', '3Stars'])

const quote = (s) => `"${String(s ?? '').replace(/"/g, '""')}"`

// Univer 的阈值 {type, value} -> Excel cfvo
function toCfvo(v) {
  if (!v) return { type: 'min' }
  const type = ['num', 'min', 'max', 'percent', 'percentile', 'formula'].includes(v.type) ? v.type : 'num'
  if (type === 'min' || type === 'max') return { type }
  if (type === 'formula') return { type, value: toExcelFormula(v.value) }
  return { type, value: Number(v.value ?? 0) }
}

function fromCfvo(c) {
  if (!c) return { type: 'min' }
  const map = { num: 'num', min: 'min', max: 'max', percent: 'percent', percentile: 'percentile', formula: 'formula', autoMin: 'min', autoMax: 'max' }
  const type = map[c.type] || 'num'
  if (type === 'min' || type === 'max') return { type }
  if (type === 'formula') return { type, value: fromExcelFormula(c.value) }
  const n = Number(c.value)
  if (Number.isNaN(n)) return { type: 'formula', value: fromExcelFormula(c.value) }
  return { type, value: n }
}

// ---------- 导出 ----------
// 把一个工作表的 Univer 条件格式规则转成 ExcelJS 的 conditionalFormattings 数组
export function univerCfToExcel(rules, { sheetName, warnings }) {
  const out = []
  rules.forEach((cf, index) => {
    const ranges = cf.ranges || []
    if (!ranges.length) return
    const ref = rangesToSqref(ranges)
    const first = ranges[0]
    const tl = cellRef(first.startRow, first.startColumn)
    const absFirst = rangeToA1(first, true)
    const priority = index + 1
    const r = cf.rule
    let rule = null

    if (r.type === 'highlightCell') {
      const style = univerHighlightToDxf(r.style)
      switch (r.subType) {
        case 'number': {
          if (!NUM_OPS.has(r.operator)) break
          const vals = Array.isArray(r.value) ? r.value : [r.value ?? 0]
          rule = { type: 'cellIs', operator: r.operator, formulae: vals.map((v) => String(v)), style }
          break
        }
        case 'text': {
          const t = quote(r.value)
          const f = {
            containsText: `NOT(ISERROR(SEARCH(${t},${tl})))`,
            notContainsText: `ISERROR(SEARCH(${t},${tl}))`,
            beginsWith: `LEFT(${tl},LEN(${t}))=${t}`,
            endsWith: `RIGHT(${tl},LEN(${t}))=${t}`,
            equal: `${tl}=${t}`,
            notEqual: `${tl}<>${t}`,
            containsBlanks: `LEN(TRIM(${tl}))=0`,
            notContainsBlanks: `LEN(TRIM(${tl}))>0`,
            containsErrors: `ISERROR(${tl})`,
            notContainsErrors: `NOT(ISERROR(${tl}))`,
          }[r.operator]
          if (f) rule = { type: 'expression', formulae: [f], style }
          break
        }
        case 'timePeriod':
          rule = { type: 'timePeriod', timePeriod: r.operator, style }
          break
        case 'rank':
          rule = { type: 'top10', rank: Number(r.value) || 10, percent: !!r.isPercent, bottom: !!r.isBottom, style }
          break
        case 'average': {
          if (r.operator === 'greaterThan' || r.operator === 'lessThan') {
            rule = { type: 'aboveAverage', aboveAverage: r.operator === 'greaterThan', style }
          } else {
            const op = { greaterThanOrEqual: '>=', lessThanOrEqual: '<=', equal: '=', notEqual: '<>' }[r.operator] || '>'
            rule = { type: 'expression', formulae: [`${tl}${op}AVERAGE(${absFirst})`], style }
          }
          break
        }
        case 'uniqueValues':
          rule = { type: 'expression', formulae: [`COUNTIF(${absFirst},${tl})=1`], style }
          break
        case 'duplicateValues':
          rule = { type: 'expression', formulae: [`COUNTIF(${absFirst},${tl})>1`], style }
          break
        case 'formula':
          rule = { type: 'expression', formulae: [toExcelFormula(r.value)], style }
          break
      }
    } else if (r.type === 'colorScale') {
      const cfg = [...(r.config || [])].sort((a, b) => a.index - b.index)
      rule = {
        type: 'colorScale',
        cfvo: cfg.map((c) => toCfvo(c.value)),
        color: cfg.map((c) => ({ argb: toArgb(c.color) || 'FF000000' })),
      }
    } else if (r.type === 'dataBar') {
      const c = r.config || {}
      rule = {
        type: 'dataBar',
        cfvo: [toCfvo(c.min), toCfvo(c.max)],
        color: { argb: toArgb(c.positiveColor) || 'FF638EC6' },
        gradient: !!c.isGradient,
        negativeFillColor: { argb: toArgb(c.nativeColor) || 'FFFF0000' },
      }
    } else if (r.type === 'iconSet') {
      const cfg = r.config || []
      const n = cfg.length
      const iconType = cfg[0]?.iconType
      if (n && iconType && iconType !== 'EMPTY_ICON_TYPE') {
        if (cfg.some((c) => c.iconType !== iconType)) warnings.push(`${sheetName}：自定义图标集导出为统一图标集 ${iconType}`)
        // Univer 从“最好”到“最差”排列，Excel 的 cfvo 从最低阈值到最高阈值排列
        const reverse = cfg.some((c, i) => Number(c.iconId) !== i)
        const cfvo = [{ type: 'percent', value: 0 }]
        for (let k = 1; k < n; k++) {
          const item = cfg[n - 1 - k]
          cfvo.push({ ...toCfvo(item.value), gte: item.operator !== 'greaterThan' })
        }
        rule = { type: 'iconSet', iconSet: iconType.replace(/^_/, ''), cfvo, showValue: r.isShowValue !== false, reverse: reverse || undefined }
        if (iconType === '_5Felling') warnings.push(`${sheetName}：表情图标集在 Excel 里没有，导出为 5Quarters`)
        if (iconType === '_5Felling') rule.iconSet = '5Quarters'
      }
    }

    if (!rule) {
      warnings.push(`${sheetName}：条件格式 ${ref} 的规则类型 ${r.type}/${r.subType || ''} 暂不支持导出`)
      return
    }
    rule.priority = priority
    if (cf.stopIfTrue) rule.stopIfTrue = true
    out.push({ ref, rules: [rule] })
  })
  return out
}

// ---------- 导入 ----------
function parseNumberOrNull(s) {
  if (s === undefined || s === null || s === '') return null
  const n = Number(String(s).trim())
  return Number.isNaN(n) ? null : n
}

function firstQuoted(formula) {
  const m = String(formula || '').match(/"((?:[^"]|"")*)"/)
  return m ? m[1].replace(/""/g, '"') : null
}

// 把常见的公式规则还原成 Univer 的专用规则（导出时文本、重复值等规则是按这些公式写的，
// 很多其他工具生成的文件也用同样的公式）
function recognizeExpression(formula, tl, style) {
  const f = String(formula).replace(/^=/, '').replace(/\s+/g, '')
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const X = esc(tl)
  const Q = '"((?:[^"]|"")*)"'
  const unq = (s) => s.replace(/""/g, '"')
  const text = (operator, value = '') => ({ type: 'highlightCell', subType: 'text', operator, value, style })
  let m
  if ((m = f.match(new RegExp(`^NOT\\(ISERROR\\(SEARCH\\(${Q},${X}\\)\\)\\)$`, 'i')))) return text('containsText', unq(m[1]))
  if ((m = f.match(new RegExp(`^ISERROR\\(SEARCH\\(${Q},${X}\\)\\)$`, 'i')))) return text('notContainsText', unq(m[1]))
  if ((m = f.match(new RegExp(`^LEFT\\(${X},LEN\\(${Q}\\)\\)=${Q}$`, 'i'))) && m[1] === m[2]) return text('beginsWith', unq(m[1]))
  if ((m = f.match(new RegExp(`^RIGHT\\(${X},LEN\\(${Q}\\)\\)=${Q}$`, 'i'))) && m[1] === m[2]) return text('endsWith', unq(m[1]))
  if ((m = f.match(new RegExp(`^${X}=${Q}$`, 'i')))) return text('equal', unq(m[1]))
  if ((m = f.match(new RegExp(`^${X}<>${Q}$`, 'i')))) return text('notEqual', unq(m[1]))
  if (new RegExp(`^LEN\\(TRIM\\(${X}\\)\\)=0$`, 'i').test(f)) return text('containsBlanks')
  if (new RegExp(`^LEN\\(TRIM\\(${X}\\)\\)>0$`, 'i').test(f)) return text('notContainsBlanks')
  if (new RegExp(`^ISERROR\\(${X}\\)$`, 'i').test(f)) return text('containsErrors')
  if (new RegExp(`^NOT\\(ISERROR\\(${X}\\)\\)$`, 'i').test(f)) return text('notContainsErrors')
  if ((m = f.match(new RegExp(`^COUNTIF\\((\\$[A-Z]+\\$\\d+(?::\\$[A-Z]+\\$\\d+)?),${X}\\)(>1|=1)$`, 'i')))) {
    return { type: 'highlightCell', subType: m[2] === '>1' ? 'duplicateValues' : 'uniqueValues', style }
  }
  if ((m = f.match(new RegExp(`^${X}(>=|<=|<>|=|>|<)AVERAGE\\(\\$[A-Z]+\\$\\d+(?::\\$[A-Z]+\\$\\d+)?\\)$`, 'i')))) {
    const operator = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '<>': 'notEqual' }[m[1]]
    return { type: 'highlightCell', subType: 'average', operator, style }
  }
  return null
}

// ExcelJS 的 conditionalFormattings -> Univer 规则数组（按优先级从高到低）
export function excelCfToUniver(conditionalFormattings, { theme, warnings, sheetName }) {
  const flat = []
  for (const cf of conditionalFormattings || []) {
    let ranges
    try { ranges = sqrefToRanges(cf.ref) } catch { continue }
    for (const rule of cf.rules || []) flat.push({ ranges, rule })
  }
  flat.sort((a, b) => (a.rule.priority ?? 0) - (b.rule.priority ?? 0))

  const result = []
  for (const { ranges, rule } of flat) {
    const first = ranges[0]
    const tl = cellRef(first.startRow, first.startColumn)
    const style = dxfToUniverHighlight(rule.style, theme)
    let r = null
    switch (rule.type) {
      case 'cellIs': {
        const vals = (rule.formulae || []).map(parseNumberOrNull)
        const op = rule.operator
        if (NUM_OPS.has(op) && vals.length && vals.every((v) => v !== null)) {
          r = { type: 'highlightCell', subType: 'number', operator: op, value: op === 'between' || op === 'notBetween' ? [vals[0], vals[1] ?? vals[0]] : vals[0], style }
        } else if ((op === 'equal' || op === 'notEqual') && /^".*"$/.test(rule.formulae?.[0] || '')) {
          r = { type: 'highlightCell', subType: 'text', operator: op, value: firstQuoted(rule.formulae[0]), style }
        } else if (NUM_OPS.has(op) && rule.formulae?.length) {
          // 阈值是公式/引用：转成“自定义公式”
          const f = rule.formulae.map((x) => fromExcelFormula(x).slice(1))
          const sym = { greaterThan: '>', greaterThanOrEqual: '>=', lessThan: '<', lessThanOrEqual: '<=', equal: '=', notEqual: '<>' }[op]
          const expr = sym ? `${tl}${sym}(${f[0]})`
            : op === 'between' ? `AND(${tl}>=(${f[0]}),${tl}<=(${f[1]}))` : `OR(${tl}<(${f[0]}),${tl}>(${f[1]}))`
          r = { type: 'highlightCell', subType: 'formula', value: '=' + expr, style }
        }
        break
      }
      case 'expression':
        if (rule.formulae?.[0]) r = recognizeExpression(rule.formulae[0], tl, style) || { type: 'highlightCell', subType: 'formula', value: fromExcelFormula(rule.formulae[0]), style }
        break
      case 'containsText': // ExcelJS 把 containsText / containsBlanks / containsErrors 等都归到这里，operator 是具体类型
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith': {
        const op = rule.type === 'containsText' ? (rule.operator || 'containsText') : rule.type
        const needsText = ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(op)
        const text = needsText ? (rule.text ?? firstQuoted(rule.formulae?.[0])) : undefined
        if (!needsText || text !== null) r = { type: 'highlightCell', subType: 'text', operator: op, value: text ?? '', style }
        else if (rule.formulae?.[0]) r = { type: 'highlightCell', subType: 'formula', value: fromExcelFormula(rule.formulae[0]), style }
        break
      }
      case 'timePeriod':
        if (rule.timePeriod) r = { type: 'highlightCell', subType: 'timePeriod', operator: rule.timePeriod, style }
        break
      case 'top10':
        r = { type: 'highlightCell', subType: 'rank', isBottom: !!rule.bottom, isPercent: !!rule.percent, value: rule.rank ?? 10, style }
        break
      case 'aboveAverage':
        r = { type: 'highlightCell', subType: 'average', operator: rule.aboveAverage === false ? 'lessThan' : 'greaterThan', style }
        break
      case 'duplicateValues':
        r = { type: 'highlightCell', subType: 'duplicateValues', style }
        break
      case 'uniqueValues':
        r = { type: 'highlightCell', subType: 'uniqueValues', style }
        break
      case 'containsBlanks':
      case 'notContainsBlanks':
      case 'containsErrors':
      case 'notContainsErrors':
        r = { type: 'highlightCell', subType: 'text', operator: rule.type, value: '', style }
        break
      case 'colorScale': {
        const colors = (rule.color || []).map((c) => excelColorToHex(c, theme) || '#000000')
        r = { type: 'colorScale', config: (rule.cfvo || []).map((c, i) => ({ index: i, color: colors[i] || colors[colors.length - 1], value: fromCfvo(c) })) }
        break
      }
      case 'dataBar': {
        const cfvo = rule.cfvo || []
        r = {
          type: 'dataBar',
          isShowValue: rule.showValue !== false,
          config: {
            min: fromCfvo(cfvo[0]),
            max: fromCfvo(cfvo[1] || { type: 'max' }),
            isGradient: rule.gradient !== false,
            positiveColor: excelColorToHex(rule.color, theme) || '#638EC6',
            nativeColor: excelColorToHex(rule.negativeFillColor, theme) || '#FF0000',
          },
        }
        break
      }
      case 'iconSet': {
        const cfvo = rule.cfvo || []
        const n = cfvo.length
        if (!n) break
        const iconType = ICON_SETS.has(rule.iconSet) ? rule.iconSet : '3TrafficLights1' // Excel 缺省图标集就是 3TrafficLights1
        const config = []
        for (let i = 0; i < n; i++) {
          const excelIdx = n - 1 - i // Univer 第 i 个（从最好开始）对应 Excel 第 n-1-i 个阈值
          const iconId = String(rule.reverse ? n - 1 - i : i)
          if (i < n - 1) {
            const c = cfvo[excelIdx]
            config.push({ operator: c.gte === false ? 'greaterThan' : 'greaterThanOrEqual', value: fromCfvo(c), iconType, iconId })
          } else {
            const prev = config[i - 1]
            config.push({ operator: prev ? OPPOSITE[prev.operator] : 'lessThan', value: prev ? { ...prev.value } : { type: 'num', value: 0 }, iconType, iconId })
          }
        }
        r = { type: 'iconSet', isShowValue: rule.showValue !== false, config }
        break
      }
    }
    if (!r) {
      warnings.push(`${sheetName}：条件格式 ${rangesToSqref(ranges)} 的 ${rule.type} 规则暂不支持导入`)
      continue
    }
    result.push({ cfId: randomId(8), ranges, stopIfTrue: !!rule.stopIfTrue, rule: r })
  }
  return result
}
