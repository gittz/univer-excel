// 直接读写 xlsx 包里的 XML，补上 ExcelJS 不支持的部分：
// - 定义名称（公式型名称、工作表级作用域、隐藏、备注）
// - 筛选条件（autoFilter 下的 filterColumn）
// - Excel 365 的“新式评论”（threadedComments + persons）
import JSZip from 'jszip'

const esc = (s) => String(s ?? '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

// ---------- 导出：修补 ExcelJS 写出的文件 ----------
function filterColumnXml(col) {
  const inner = []
  if (col.filters) {
    const attrs = col.filters.blank ? ' blank="1"' : ''
    const items = (col.filters.filters || []).map((v) => `<filter val="${esc(v)}"/>`).join('')
    inner.push(`<filters${attrs}>${items}</filters>`)
  } else if (col.customFilters?.customFilters?.length) {
    const and = col.customFilters.and ? ' and="1"' : ''
    const items = col.customFilters.customFilters.map((c) => {
      const op = c.operator && c.operator !== 'equal' ? ` operator="${esc(c.operator)}"` : ''
      return `<customFilter${op} val="${esc(c.val)}"/>`
    }).join('')
    inner.push(`<customFilters${and}>${items}</customFilters>`)
  } else {
    return ''
  }
  return `<filterColumn colId="${col.colId}">${inner.join('')}</filterColumn>`
}

/**
 * @param {ArrayBuffer} buffer ExcelJS 写出的 xlsx
 * @param {{ definedNames: Array<{name:string, formula:string, localSheetId?:number, hidden?:boolean, comment?:string}>,
 *           filters: Record<number, {ref:string, columns:any[]}> }} patch
 */
export async function patchXlsxExport(buffer, { definedNames = [], filters = {}, validations = {}, unlockAllColumns = [] }) {
  const zip = await JSZip.loadAsync(buffer)

  // 定义名称
  const wbPath = 'xl/workbook.xml'
  let wbXml = await zip.file(wbPath).async('string')
  wbXml = wbXml.replace(/<definedNames>[\s\S]*?<\/definedNames>|<definedNames\/>/g, '')
  if (definedNames.length) {
    const items = definedNames.map((d) => {
      const attrs = [`name="${esc(d.name)}"`]
      if (d.localSheetId !== undefined && d.localSheetId !== null) attrs.push(`localSheetId="${d.localSheetId}"`)
      if (d.hidden) attrs.push('hidden="1"')
      if (d.comment) attrs.push(`comment="${esc(d.comment)}"`)
      return `<definedName ${attrs.join(' ')}>${esc(d.formula)}</definedName>`
    }).join('')
    wbXml = wbXml.replace('</sheets>', `</sheets><definedNames>${items}</definedNames>`)
  }
  zip.file(wbPath, wbXml)

  // ExcelJS 会在数据条规则里写一个空的 <x14:id/> 扩展，但不写对应的 x14 定义，Excel 会提示文件损坏；去掉它，保留标准数据条
  for (const path of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml = await zip.file(path).async('string')
    const fixed = xml.replace(/(<cfRule[^>]*type="dataBar"[^>]*>[\s\S]*?<\/dataBar>)<extLst>[\s\S]*?<\/extLst>/g, '$1')
    if (fixed !== xml) zip.file(path, fixed)
  }

  // 筛选条件
  for (const [index, f] of Object.entries(filters)) {
    const path = `xl/worksheets/sheet${Number(index) + 1}.xml`
    const file = zip.file(path)
    if (!file) continue
    let xml = await file.async('string')
    const cols = (f.columns || []).map(filterColumnXml).join('')
    if (!cols) continue
    xml = xml.replace(/<autoFilter ref="([^"]+)"\s*\/>/, (_all, ref) => `<autoFilter ref="${ref}">${cols}</autoFilter>`)
    zip.file(path, xml)
  }

  // 数据验证
  for (const [index, list] of Object.entries(validations)) {
    if (!list?.length) continue
    const path = `xl/worksheets/sheet${Number(index) + 1}.xml`
    const file = zip.file(path)
    if (!file) continue
    let xml = await file.async('string')
    xml = xml.replace(/<dataValidations[\s\S]*?<\/dataValidations>/, '')
    const items = list.map(({ sqref, v }) => {
      const attrs = [`type="${v.type}"`]
      if (v.allowBlank) attrs.push('allowBlank="1"')
      if (v.showInputMessage) attrs.push('showInputMessage="1"')
      if (v.showErrorMessage) attrs.push('showErrorMessage="1"')
      if (v.errorStyle && v.errorStyle !== 'stop') attrs.push(`errorStyle="${v.errorStyle}"`)
      if (v.operator && v.operator !== 'between' && !['list', 'custom'].includes(v.type)) attrs.push(`operator="${v.operator}"`)
      if (v.errorTitle) attrs.push(`errorTitle="${esc(v.errorTitle)}"`)
      if (v.error) attrs.push(`error="${esc(v.error)}"`)
      if (v.promptTitle) attrs.push(`promptTitle="${esc(v.promptTitle)}"`)
      if (v.prompt) attrs.push(`prompt="${esc(v.prompt)}"`)
      attrs.push(`sqref="${sqref}"`)
      const f = (v.formulae || []).slice(0, 2).map((x, i) => `<formula${i + 1}>${esc(x)}</formula${i + 1}>`).join('')
      return `<dataValidation ${attrs.join(' ')}>${f}</dataValidation>`
    }).join('')
    const block = `<dataValidations count="${list.length}">${items}</dataValidations>`
    // OOXML 要求 dataValidations 出现在这些元素之前
    const after = ['hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst']
    // 只在顶层元素里找插入点：从最后一个条件格式 / sheetData 之后开始搜索，避免匹配到 cfRule 里的 extLst
    const from = Math.max(xml.lastIndexOf('</conditionalFormatting>'), xml.lastIndexOf('</sheetData>'), xml.lastIndexOf('</mergeCells>'))
    let pos = -1
    for (const tag of after) {
      const m = xml.slice(from).search(new RegExp(`<${tag}[\\s>/]`))
      if (m >= 0 && (pos < 0 || from + m < pos)) pos = from + m
    }
    if (pos < 0) pos = xml.lastIndexOf('</worksheet>')
    xml = xml.slice(0, pos) + block + xml.slice(pos)
    zip.file(path, xml)
  }

  // 整张表的列默认“未锁定”：新增一个未锁定的单元格样式，给 1~16384 列都挂上（保留原有列宽）
  if (unlockAllColumns.length) {
    let styles = await zip.file('xl/styles.xml').async('string')
    const m = styles.match(/<cellXfs count="(\d+)">/)
    const xfIndex = Number(m[1])
    styles = styles.replace(m[0], `<cellXfs count="${xfIndex + 1}">`)
      .replace('</cellXfs>', '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf></cellXfs>')
    zip.file('xl/styles.xml', styles)
    for (const index of unlockAllColumns) {
      const path = `xl/worksheets/sheet${Number(index) + 1}.xml`
      const file = zip.file(path)
      if (!file) continue
      let xml = await file.async('string')
      const defaultWidth = (xml.match(/<sheetFormatPr[^>]*defaultColWidth="([\d.]+)"/) || [])[1] || '8.43'
      const existing = []
      const colsMatch = xml.match(/<cols>([\s\S]*?)<\/cols>/)
      if (colsMatch) {
        for (const c of colsMatch[1].match(/<col [^>]*\/>/g) || []) {
          existing.push({ min: Number(c.match(/min="(\d+)"/)[1]), max: Number(c.match(/max="(\d+)"/)[1]), tag: c })
        }
      }
      existing.sort((a, b) => a.min - b.min)
      const out = []
      let next = 1
      const gap = (from, to) => { if (from <= to) out.push(`<col min="${from}" max="${to}" width="${defaultWidth}" style="${xfIndex}"/>`) }
      for (const c of existing) {
        gap(next, c.min - 1)
        out.push(/ style="/.test(c.tag) ? c.tag : c.tag.replace('<col ', `<col style="${xfIndex}" `))
        next = c.max + 1
      }
      gap(next, 16384)
      const cols = `<cols>${out.join('')}</cols>`
      xml = colsMatch ? xml.replace(colsMatch[0], cols) : xml.replace('<sheetData', `${cols}<sheetData`)
      zip.file(path, xml)
    }
  }

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
}

// ---------- 导入：读取 ExcelJS 读不到的部分 ----------
function parseXml(text) {
  return new DOMParser().parseFromString(text, 'application/xml')
}

const byTag = (node, tag) => Array.from(node.getElementsByTagName(tag))

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1)
  const parts = base.split('/').slice(0, -1)
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.') parts.push(seg)
  }
  return parts.join('/')
}

async function readRels(zip, partPath) {
  const dir = partPath.split('/').slice(0, -1).join('/')
  const name = partPath.split('/').pop()
  const relsPath = `${dir}/_rels/${name}.rels`
  const file = zip.file(relsPath)
  if (!file) return []
  const doc = parseXml(await file.async('string'))
  return byTag(doc, 'Relationship').map((r) => ({
    id: r.getAttribute('Id'),
    type: r.getAttribute('Type') || '',
    target: resolvePath(partPath, r.getAttribute('Target') || ''),
  }))
}

function parseFilterColumn(el) {
  const colId = Number(el.getAttribute('colId'))
  const filters = byTag(el, 'filters')[0]
  if (filters) {
    const out = { colId, filters: {} }
    if (filters.getAttribute('blank') === '1' || filters.getAttribute('blank') === 'true') out.filters.blank = true
    const vals = byTag(filters, 'filter').map((f) => f.getAttribute('val'))
    if (vals.length) out.filters.filters = vals
    return out
  }
  const custom = byTag(el, 'customFilters')[0]
  if (custom) {
    const list = byTag(custom, 'customFilter').map((c) => {
      const val = c.getAttribute('val') ?? ''
      const op = c.getAttribute('operator')
      const num = Number(val)
      return { val: val !== '' && !Number.isNaN(num) ? num : val, ...(op && op !== 'equal' ? { operator: op } : {}) }
    })
    return { colId, customFilters: { ...(custom.getAttribute('and') === '1' || custom.getAttribute('and') === 'true' ? { and: 1 } : {}), customFilters: list.slice(0, 2) } }
  }
  return { colId, unsupported: el.firstElementChild?.tagName || 'unknown' }
}

// 有些工具（openpyxl、部分国产软件等）在 .rels 里写绝对路径（Target="/xl/drawings/drawing1.xml"），
// ExcelJS 读这种文件会崩溃。这里统一改写成相对路径后再交给 ExcelJS。
export async function normalizeXlsxForExcelJS(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer)
  let changed = false
  for (const relsPath of Object.keys(zip.files).filter((n) => n.endsWith('.rels'))) {
    const xml = await zip.file(relsPath).async('string')
    // 'xl/worksheets/_rels/sheet1.xml.rels' 描述的是 'xl/worksheets/sheet1.xml'
    const baseDir = relsPath.replace(/_rels\/[^/]+\.rels$/, '').replace(/\/$/, '')
    const fixed = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      if (/TargetMode="External"/.test(tag)) return tag
      return tag.replace(/Target="\/([^"]*)"/, (_m, abs) => {
        const from = baseDir ? baseDir.split('/') : []
        const to = abs.split('/')
        let i = 0
        while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++
        const rel = [...Array(from.length - i).fill('..'), ...to.slice(i)].join('/')
        return `Target="${rel}"`
      })
    })
    if (fixed !== xml) { zip.file(relsPath, fixed); changed = true }
  }

  // 绘图 XML 没有 xdr: 前缀（默认命名空间写法）时，ExcelJS 解析不出图片，补上前缀
  for (const path of Object.keys(zip.files).filter((n) => /^xl\/drawings\/[^/]+\.xml$/.test(n))) {
    const xml = await zip.file(path).async('string')
    if (!/<wsDr[\s>]/.test(xml)) continue
    const DRAW_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    const fixed = xml
      .replace(`xmlns="${DRAW_NS}"`, `xmlns:xdr="${DRAW_NS}"`)
      .replace(/<(\/?)([A-Za-z][\w.-]*)(?=[\s>/])/g, (m, slash, tag) => `<${slash}xdr:${tag}`)
    zip.file(path, fixed)
    changed = true
  }

  // ExcelJS 只认 xl/commentsN.xml 和 xl/drawings/vmlDrawingN.vml 这种文件名，其他命名的批注文件改名
  let seq = 1000
  for (const relsPath of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(n))) {
    let xml = await zip.file(relsPath).async('string')
    const sheetPart = relsPath.replace('_rels/', '').replace(/\.rels$/, '')
    let touched = false
    xml = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const type = (tag.match(/Type="([^"]*)"/) || [])[1] || ''
      const target = (tag.match(/Target="([^"]*)"/) || [])[1]
      if (!target || /TargetMode="External"/.test(tag)) return tag
      const abs = resolvePath(sheetPart, target)
      let newAbs = null
      if (/\/comments$/.test(type) && !/^xl\/comments\d+\.xml$/.test(abs)) newAbs = `xl/comments${seq++}.xml`
      if (/\/vmlDrawing$/.test(type) && !/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(abs)) newAbs = `xl/drawings/vmlDrawing${seq++}.vml`
      if (!newAbs || !zip.file(abs)) return tag
      zip.file(newAbs, zip.file(abs).async('uint8array'))
      zip.remove(abs)
      touched = true
      const rel = newAbs.startsWith('xl/drawings/') ? `../drawings/${newAbs.split('/').pop()}` : `../${newAbs.split('/').pop()}`
      return tag.replace(/Target="[^"]*"/, `Target="${rel}"`)
    })
    if (touched) { zip.file(relsPath, xml); changed = true }
  }
  return changed ? zip.generateAsync({ type: 'arraybuffer' }) : arrayBuffer
}

export async function readXlsxExtras(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer)
  const result = { definedNames: [], sheets: [] }
  const wbPath = 'xl/workbook.xml'
  const wbFile = zip.file(wbPath)
  if (!wbFile) return result
  const wbDoc = parseXml(await wbFile.async('string'))
  const wbRels = await readRels(zip, wbPath)

  for (const d of byTag(wbDoc, 'definedName')) {
    result.definedNames.push({
      name: d.getAttribute('name'),
      formula: d.textContent,
      localSheetId: d.hasAttribute('localSheetId') ? Number(d.getAttribute('localSheetId')) : undefined,
      hidden: d.getAttribute('hidden') === '1' || d.getAttribute('hidden') === 'true',
      comment: d.getAttribute('comment') || undefined,
    })
  }

  // 人员（新式评论的作者）
  const persons = {}
  for (const rel of wbRels.filter((r) => /\/person$/.test(r.type))) {
    const f = zip.file(rel.target)
    if (!f) continue
    for (const p of byTag(parseXml(await f.async('string')), 'person')) persons[p.getAttribute('id')] = p.getAttribute('displayName') || ''
  }

  const sheetEls = byTag(wbDoc, 'sheet')
  for (const el of sheetEls) {
    const rid = el.getAttribute('r:id') || el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
    const rel = wbRels.find((r) => r.id === rid)
    const info = { name: el.getAttribute('name'), autoFilter: null, threadedComments: [], notes: {}, validations: [] }
    result.sheets.push(info)
    if (!rel || !zip.file(rel.target)) continue
    const sheetDoc = parseXml(await zip.file(rel.target).async('string'))

    // 工作表级 autoFilter（表格里的 autoFilter 在 table 文件里，不在这里）
    const af = byTag(sheetDoc, 'autoFilter').find((n) => n.parentNode === sheetDoc.documentElement)
    if (af) info.autoFilter = { ref: af.getAttribute('ref'), columns: byTag(af, 'filterColumn').map(parseFilterColumn) }

    // 数据验证：普通的 <dataValidation>，以及 Excel 放在扩展区里的 <x14:dataValidation>（引用其他工作表的下拉列表就存在这里）
    for (const dv of byTag(sheetDoc, 'dataValidation')) {
      info.validations.push({
        sqref: dv.getAttribute('sqref') || '',
        type: dv.getAttribute('type') || 'none',
        operator: dv.getAttribute('operator') || undefined,
        allowBlank: ['1', 'true'].includes(dv.getAttribute('allowBlank')),
        showInputMessage: ['1', 'true'].includes(dv.getAttribute('showInputMessage')),
        showErrorMessage: ['1', 'true'].includes(dv.getAttribute('showErrorMessage')),
        errorStyle: dv.getAttribute('errorStyle') || undefined,
        error: dv.getAttribute('error') || undefined,
        errorTitle: dv.getAttribute('errorTitle') || undefined,
        prompt: dv.getAttribute('prompt') || undefined,
        promptTitle: dv.getAttribute('promptTitle') || undefined,
        formulae: [byTag(dv, 'formula1')[0], byTag(dv, 'formula2')[0]].filter(Boolean).map((f) => f.textContent),
      })
    }
    for (const dv of byTag(sheetDoc, 'x14:dataValidation')) {
      const f = (tag) => { const n = byTag(dv, tag)[0]; return n ? (byTag(n, 'xm:f')[0]?.textContent ?? n.textContent) : undefined }
      info.validations.push({
        sqref: byTag(dv, 'xm:sqref')[0]?.textContent || '',
        type: dv.getAttribute('type') || 'none',
        operator: dv.getAttribute('operator') || undefined,
        allowBlank: ['1', 'true'].includes(dv.getAttribute('allowBlank')),
        showInputMessage: ['1', 'true'].includes(dv.getAttribute('showInputMessage')),
        showErrorMessage: ['1', 'true'].includes(dv.getAttribute('showErrorMessage')),
        errorStyle: dv.getAttribute('errorStyle') || undefined,
        error: dv.getAttribute('error') || undefined,
        errorTitle: dv.getAttribute('errorTitle') || undefined,
        prompt: dv.getAttribute('prompt') || undefined,
        promptTitle: dv.getAttribute('promptTitle') || undefined,
        formulae: [f('x14:formula1'), f('x14:formula2')].filter((x) => x !== undefined),
      })
    }

    // 新式评论
    const sheetRels = await readRels(zip, rel.target)
    // 旧式批注（note）：直接读 XML，ExcelJS 对没有格式片段的批注会读成空文本
    for (const cr of sheetRels.filter((r) => /\/comments$/.test(r.type))) {
      const f = zip.file(cr.target)
      if (!f) continue
      const doc = parseXml(await f.async('string'))
      const authors = byTag(doc, 'author').map((a) => a.textContent)
      for (const c of byTag(doc, 'comment')) {
        const text = byTag(c, 't').map((t) => t.textContent).join('')
        info.notes[c.getAttribute('ref')] = { text, author: authors[Number(c.getAttribute('authorId'))] || '' }
      }
    }
    for (const tr of sheetRels.filter((r) => /\/threadedComment$/.test(r.type))) {
      const f = zip.file(tr.target)
      if (!f) continue
      for (const c of byTag(parseXml(await f.async('string')), 'threadedComment')) {
        info.threadedComments.push({
          id: c.getAttribute('id'),
          ref: c.getAttribute('ref'),
          dT: c.getAttribute('dT'),
          personId: c.getAttribute('personId'),
          author: persons[c.getAttribute('personId')] || '',
          parentId: c.getAttribute('parentId') || undefined,
          done: c.getAttribute('done') === '1',
          text: byTag(c, 'text')[0]?.textContent || '',
        })
      }
    }
  }
  return result
}
