// 数据验证：Univer 规则 <-> Excel dataValidations
import { rangeToA1, sqrefToRanges, clampRange, formatDate, serialToDate, dateToSerial, randomId } from './utils.js'
import { toExcelFormula, fromExcelFormula } from './formula.js'

const ERROR_STYLE_TO_EXCEL = { 0: 'information', 1: 'stop', 2: 'warning' }
const ERROR_STYLE_TO_UNIVER = { information: 0, stop: 1, warning: 2 }

function parseListItems(formula1) {
  const s = String(formula1 ?? '')
  if (s.startsWith('[')) {
    try { const arr = JSON.parse(s); if (Array.isArray(arr)) return arr.map(String) } catch { /* fallthrough */ }
  }
  return s.split(',').map((x) => x.trim()).filter((x) => x !== '')
}

// '=[unitId]Sheet1!A1:A3' -> 'Sheet1!$A$1:$A$3' 之类的 Excel 引用
function toExcelRef(formula) {
  return toExcelFormula(String(formula).replace(/\[[^\]]*\]/g, ''))
}

function dateFormulaToExcel(v) {
  if (v === undefined || v === null || v === '') return undefined
  const s = String(v)
  if (s.startsWith('=')) return toExcelFormula(s)
  const d = new Date(s.length <= 10 ? `${s}T00:00:00Z` : s)
  if (Number.isNaN(d.getTime())) return s
  return String(Math.round(dateToSerial(d) * 1e6) / 1e6)
}

function timeFormulaToExcel(v) {
  if (v === undefined || v === null || v === '') return undefined
  const s = String(v)
  if (s.startsWith('=')) return toExcelFormula(s)
  const m = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/)
  if (!m) return s
  return String((Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0)) / 86400)
}

const plain = (v) => (v === undefined || v === null || v === '' ? undefined : String(v).startsWith('=') ? toExcelFormula(v) : String(v))

// ---------- 导出 ----------
// 返回 [{ sqref, v }]，由 xml-parts 直接写进工作表 XML（ExcelJS 自己写会产生重叠区域，Excel 会提示修复）
export function univerDvToExcel(rules, { sheetName, maxRow, maxCol, warnings }) {
  const out = []
  for (const rule of rules || []) {
    const v = {}
    switch (rule.type) {
      case 'list':
      case 'listMultiple': {
        v.type = 'list'
        const f1 = String(rule.formula1 ?? '')
        if (f1.startsWith('=')) {
          v.formulae = [toExcelRef(f1)]
        } else {
          const items = parseListItems(f1)
          if (items.some((x) => x.includes(','))) warnings.push(`${sheetName}：下拉选项里含逗号，Excel 会把它拆成多项`)
          const joined = items.join(',')
          if (joined.length > 255) warnings.push(`${sheetName}：下拉选项总长度超过 Excel 的 255 字符限制`)
          v.formulae = [`"${joined.replace(/"/g, '""')}"`]
        }
        if (rule.type === 'listMultiple') warnings.push(`${sheetName}：多选下拉在 Excel 里会变成单选下拉`)
        break
      }
      case 'checkbox': {
        const on = rule.formula1 ?? 'TRUE'
        const off = rule.formula2 ?? 'FALSE'
        v.type = 'list'
        v.formulae = [`"${on},${off}"`]
        warnings.push(`${sheetName}：复选框验证在 Excel 里没有对应类型，导出为「${on}/${off}」下拉`)
        break
      }
      case 'decimal':
      case 'whole':
      case 'textLength':
        v.type = rule.type
        v.operator = rule.operator || 'between'
        v.formulae = [plain(rule.formula1), plain(rule.formula2)].filter((x) => x !== undefined)
        break
      case 'date':
        v.type = 'date'
        v.operator = rule.operator || 'between'
        v.formulae = [dateFormulaToExcel(rule.formula1), dateFormulaToExcel(rule.formula2)].filter((x) => x !== undefined)
        break
      case 'time':
        v.type = 'time'
        v.operator = rule.operator || 'between'
        v.formulae = [timeFormulaToExcel(rule.formula1), timeFormulaToExcel(rule.formula2)].filter((x) => x !== undefined)
        break
      case 'custom':
        v.type = 'custom'
        v.formulae = [toExcelFormula(rule.formula1 || '=TRUE')]
        break
      default:
        continue
    }
    if (!v.formulae?.length) continue
    // 区间运算符至少要两个值
    if ((v.operator === 'between' || v.operator === 'notBetween') && v.formulae.length === 1) v.formulae.push(v.formulae[0])

    v.allowBlank = rule.allowBlank !== false
    v.showErrorMessage = rule.showErrorMessage !== false
    v.errorStyle = ERROR_STYLE_TO_EXCEL[rule.errorStyle ?? 1] || 'stop'
    if (rule.error) v.error = rule.error
    if (rule.errorTitle) v.errorTitle = rule.errorTitle
    if (rule.showInputMessage || rule.prompt) v.showInputMessage = true
    if (rule.prompt) v.prompt = rule.prompt
    if (rule.promptTitle) v.promptTitle = rule.promptTitle

    const sqref = (rule.ranges || []).map((range) => rangeToA1(clampRange(range, maxRow, maxCol))).join(' ')
    if (sqref) out.push({ sqref, v })
  }
  return out
}

// ---------- 导入 ----------
function dateFromExcel(f) {
  if (f === undefined || f === null) return undefined
  if (f instanceof Date) return Number.isNaN(f.getTime()) ? undefined : formatDate(f)
  const n = Number(f)
  if (!Number.isNaN(n)) return formatDate(serialToDate(n))
  return fromExcelFormula(f)
}

function timeFromExcel(f) {
  if (f === undefined || f === null) return undefined
  const n = Number(f)
  if (Number.isNaN(n)) return fromExcelFormula(f)
  const secs = Math.round((n % 1) * 86400)
  const hh = String(Math.floor(secs / 3600)).padStart(2, '0')
  const mm = String(Math.floor((secs % 3600) / 60)).padStart(2, '0')
  const ss = String(secs % 60).padStart(2, '0')
  return `${hh}:${mm}:${ss}`
}

const numFromExcel = (f) => {
  if (f === undefined || f === null) return undefined
  return Number.isNaN(Number(f)) ? fromExcelFormula(f) : String(f)
}

// 从 XML 读到的数据验证列表 -> Univer 规则
export function excelDvToUniver(list, { warnings, sheetName }) {
  const rules = []
  for (const v of list || []) {
    if (!v.type || v.type === 'none' || !v.sqref) continue
    let ranges
    try { ranges = sqrefToRanges(v.sqref) } catch { continue }
    const rule = { uid: randomId(16), ranges }
    const f = v.formulae || []
    switch (v.type) {
      case 'list': {
        rule.type = 'list'
        rule.showDropDown = true
        const f1 = String(f[0] ?? '')
        if (/^".*"$/.test(f1)) rule.formula1 = JSON.stringify(f1.slice(1, -1).replace(/""/g, '"').split(','))
        else rule.formula1 = fromExcelFormula(f1)
        break
      }
      case 'whole':
      case 'decimal':
      case 'textLength':
        rule.type = v.type
        rule.operator = v.operator || 'between'
        rule.formula1 = numFromExcel(f[0])
        rule.formula2 = numFromExcel(f[1])
        break
      case 'date':
        rule.type = 'date'
        rule.operator = v.operator || 'between'
        rule.formula1 = dateFromExcel(f[0])
        rule.formula2 = dateFromExcel(f[1])
        break
      case 'time':
        rule.type = 'time'
        rule.operator = v.operator || 'between'
        rule.formula1 = timeFromExcel(f[0])
        rule.formula2 = timeFromExcel(f[1])
        break
      case 'custom':
        rule.type = 'custom'
        rule.formula1 = fromExcelFormula(f[0] ?? 'TRUE')
        break
      default:
        warnings.push(`${sheetName}：数据验证类型 ${v.type} 暂不支持导入`)
        continue
    }
    if (rule.formula2 === undefined) delete rule.formula2
    // Excel 从不把空单元格判为无效；Univer 在 allowBlank=false 时会给空单元格打红角，这里按 Excel 的表现处理
    rule.allowBlank = true
    rule.showErrorMessage = !!v.showErrorMessage
    rule.errorStyle = ERROR_STYLE_TO_UNIVER[v.errorStyle || 'stop'] ?? 1
    if (v.error) rule.error = v.error
    if (v.errorTitle) rule.errorTitle = v.errorTitle
    if (v.showInputMessage) rule.showInputMessage = true
    if (v.prompt) rule.prompt = v.prompt
    if (v.promptTitle) rule.promptTitle = v.promptTitle
    rules.push(rule)
  }
  return rules
}
