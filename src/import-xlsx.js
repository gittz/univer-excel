// 纯前端把 .xlsx 导入为 Univer 工作簿（只用开源的 ExcelJS / JSZip，不依赖 Univer Pro 或服务端）
import ExcelJS from 'exceljs'
import {
  parseThemeColors, excelCellStyleToUniver, excelFontToUniver, excelColorToHex, createSheetMetrics,
  a1ToRange, cellsToRanges, complementRects, colWidthToPx, ptToPx, dateToSerial, randomId, createCellDoc, bytesToBase64, EMU_PER_PX,
} from './xlsx/utils.js'
import { fromExcelFormula } from './xlsx/formula.js'
import { excelCfToUniver } from './xlsx/conditional-format.js'
import { excelDvToUniver } from './xlsx/data-validation.js'
import { readXlsxExtras, normalizeXlsxForExcelJS } from './xlsx/xml-parts.js'
import { withProtectionHookSuppressed, applyExcelRangePoints, applyExcelWorksheetPoints } from './protection.js'

function plainValue(v) {
  if (v === null || v === undefined) return null
  if (v instanceof Date) return { v: dateToSerial(v), t: 2 }
  if (typeof v === 'number') return { v, t: 2 }
  if (typeof v === 'boolean') return { v: v ? 1 : 0, t: 3 }
  if (typeof v === 'string') return { v, t: 1 }
  if (v.richText) return { v: v.richText.map((r) => r.text).join(''), t: 1 }
  if (v.error) return { v: v.error, t: 1 }
  if (v.text !== undefined) return plainValue(v.text)
  return null
}

function richTextToDoc(richText, theme) {
  let text = ''
  const textRuns = []
  for (const piece of richText) {
    const t = String(piece.text ?? '').replace(/\r?\n/g, '\r')
    const ts = excelFontToUniver(piece.font, theme, { keepBlack: true })
    if (t && Object.keys(ts).length) textRuns.push({ st: text.length, ed: text.length + t.length, ts })
    text += t
  }
  return createCellDoc(text.replace(/\r/g, '\n'), { textRuns })
}

const noteText = (note) => {
  if (!note) return ''
  if (typeof note === 'string') return note
  return (note.texts || []).map((t) => t.text).join('')
}

const toUniverTime = (iso) => {
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return iso || ''
  const p = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
}

const commentBody = (text) => ({
  dataStream: `${String(text).replace(/\r?\n/g, '\r')}\r\n`,
  textRuns: [],
  paragraphs: [{ startIndex: String(text).replace(/\r?\n/g, '\r').length }],
  customRanges: [],
  customBlocks: [],
  sectionBreaks: [],
  customDecorations: [],
})

export async function xlsxToWorkbookData(input, name = 'Workbook') {
  const arrayBuffer = await normalizeXlsxForExcelJS(input)
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(arrayBuffer)
  const extras = await readXlsxExtras(arrayBuffer)
  const theme = parseThemeColors(wb)
  const warnings = []
  const unitId = `workbook-${randomId(8)}`

  const styles = {}
  const styleKeys = new Map()
  const styleId = (style) => {
    const key = JSON.stringify(style)
    let id = styleKeys.get(key)
    if (!id) {
      id = 's' + styleKeys.size
      styleKeys.set(key, id)
      styles[id] = style
    }
    return id
  }

  const worksheets = wb.worksheets
  const sheetIds = worksheets.map((_, i) => `sheet-${i + 1}`)
  const sheetIdByName = Object.fromEntries(worksheets.map((ws, i) => [ws.name, sheetIds[i]]))

  const sheets = {}
  const resources = { cf: {}, dv: {}, filter: {}, note: {}, comment: {}, drawing: {}, definedName: {} }
  const post = { tables: [], protectedSheets: [], rangeProtections: [] }

  for (const [index, ws] of worksheets.entries()) {
    const id = sheetIds[index]
    const extra = extras.sheets[index] || {}
    const cellData = {}
    let maxRow = 0
    let maxCol = 0
    const covered = new Set() // 数组公式覆盖的单元格（由主单元格溢出，不单独写值）
    const threadedCells = new Set((extra.threadedComments || []).map((c) => c.ref))
    const notes = {}
    const unlockedCells = []
    const lockedCells = []

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const r = rowNumber - 1
        const c = colNumber - 1
        if (cell.protection && cell.protection.locked === false) unlockedCells.push({ row: r, col: c })
        else lockedCells.push({ row: r, col: c })
        const isSlave = cell.isMerged && cell.master && cell.master !== cell
        let data = {}

        if (!isSlave && !covered.has(`${r},${c}`)) {
          const value = cell.value
          if (cell.type === ExcelJS.ValueType.Formula) {
            data.f = fromExcelFormula(cell.formula)
            const res = plainValue(cell.result)
            if (res) Object.assign(data, res)
            if (value?.shareType === 'array' && value.ref) {
              const rg = a1ToRange(value.ref)
              for (let i = rg.startRow; i <= rg.endRow; i++) for (let j = rg.startColumn; j <= rg.endColumn; j++) if (i !== r || j !== c) covered.add(`${i},${j}`)
            }
          } else if (value && typeof value === 'object' && value.hyperlink !== undefined) {
            const text = typeof value.text === 'object' && value.text?.richText ? value.text.richText.map((x) => x.text).join('') : String(value.text ?? value.hyperlink)
            let url = value.hyperlink
            if (url.startsWith('#')) {
              // 内部链接 '#Sheet2!A1' / "#'工作表 2'!A1"
              const m = url.slice(1).match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/)
              if (m) {
                const sheetName = (m[1] || m[2]).replace(/''/g, "'")
                const gid = sheetIdByName[sheetName]
                url = gid ? `#gid=${gid}&range=${m[3].replace(/\$/g, '')}` : url
              }
            }
            data.p = createCellDoc(text, { customRanges: [{ startIndex: 0, endIndex: text.length - 1, rangeType: 0, rangeId: randomId(10), properties: { url, ...(value.tooltip ? { tooltip: value.tooltip } : {}) } }] })
            data.v = text
            data.t = 1
          } else if (value && typeof value === 'object' && value.richText) {
            data.p = richTextToDoc(value.richText, theme)
            data.v = value.richText.map((x) => x.text).join('')
            data.t = 1
          } else {
            const pv = plainValue(value)
            if (pv) Object.assign(data, pv)
          }
        }

        if ((cell.note || extra.notes?.[cell.address]) && !threadedCells.has(cell.address)) {
          const text = extra.notes?.[cell.address]?.text || noteText(cell.note)
          if (text) (notes[r] ||= {})[c] = { id: randomId(6), row: r, col: c, note: text, width: 160, height: 72, show: false }
        }

        const style = excelCellStyleToUniver(cell, theme)
        if (style) data.s = styleId(style)
        if (!Object.keys(data).length) return
        ;(cellData[r] ||= {})[c] = data
        maxRow = Math.max(maxRow, r)
        maxCol = Math.max(maxCol, c)
      })
    })
    // 空单元格上的批注（ExcelJS 可能没有为它建单元格）
    for (const [ref, n] of Object.entries(extra.notes || {})) {
      if (threadedCells.has(ref) || !n.text) continue
      const rg = a1ToRange(ref)
      if (notes[rg.startRow]?.[rg.startColumn]) continue
      ;(notes[rg.startRow] ||= {})[rg.startColumn] = { id: randomId(6), row: rg.startRow, col: rg.startColumn, note: n.text, width: 160, height: 72, show: false }
    }
    // 数组公式覆盖区域里 ExcelJS 已经写入的缓存值要去掉，否则 Univer 会认为溢出区域被占用
    for (const key of covered) {
      const [r, c] = key.split(',').map(Number)
      const d = cellData[r]?.[c]
      if (d) {
        delete d.v
        delete d.t
        delete d.p
        if (!Object.keys(d).length) delete cellData[r][c]
      }
    }

    const columnData = {}
    ;(ws.columns || []).forEach((col, i) => {
      if (!col) return
      const d = {}
      if (col.width && col.isCustomWidth !== false) d.w = colWidthToPx(col.width)
      if (col.hidden) d.hd = 1
      if (Object.keys(d).length) columnData[i] = d
    })

    const rowData = {}
    const hiddenRows = []
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.findRow(r)
      if (!row) continue
      const d = {}
      if (row.height) d.h = ptToPx(row.height)
      if (row.hidden) { d.hd = 1; hiddenRows.push(r - 1) }
      if (Object.keys(d).length) rowData[r - 1] = d
    }

    const mergeData = (ws.model.merges || []).map(a1ToRange)
    for (const m of mergeData) {
      maxRow = Math.max(maxRow, m.endRow)
      maxCol = Math.max(maxCol, m.endColumn)
    }

    const view = (ws.views || [])[0]
    const xSplit = view?.state === 'frozen' ? view.xSplit || 0 : 0
    const ySplit = view?.state === 'frozen' ? view.ySplit || 0 : 0

    const sheet = {
      id,
      name: ws.name,
      rowCount: Math.max(1000, maxRow + 100),
      columnCount: Math.max(20, maxCol + 10),
      cellData,
      mergeData,
      rowData,
      columnData,
      freeze: { xSplit, ySplit, startRow: ySplit, startColumn: xSplit },
      hidden: ws.state === 'hidden' || ws.state === 'veryHidden' ? 1 : 0,
      showGridlines: view?.showGridLines === false ? 0 : 1,
      zoomRatio: view?.zoomScale ? view.zoomScale / 100 : 1,
    }
    if (ws.properties?.defaultRowHeight) sheet.defaultRowHeight = ptToPx(ws.properties.defaultRowHeight)
    if (ws.properties?.defaultColWidth) sheet.defaultColumnWidth = colWidthToPx(ws.properties.defaultColWidth)
    const tab = excelColorToHex(ws.properties?.tabColor, theme)
    if (tab) sheet.tabColor = tab
    sheets[id] = sheet

    // 条件格式
    const cf = excelCfToUniver(ws.conditionalFormattings, { theme, warnings, sheetName: ws.name })
    if (cf.length) resources.cf[id] = cf

    // 数据验证
    const dv = excelDvToUniver(extra.validations, { warnings, sheetName: ws.name })
    if (dv.length) resources.dv[id] = dv

    // 筛选
    const afRef = extra.autoFilter?.ref || (typeof ws.autoFilter === 'string' ? ws.autoFilter : null)
    if (afRef) {
      const ref = a1ToRange(afRef)
      const filterColumns = []
      for (const col of extra.autoFilter?.columns || []) {
        if (col.unsupported) { warnings.push(`${ws.name}：筛选条件 ${col.unsupported} 暂不支持导入`); continue }
        filterColumns.push({ ...col, colId: col.colId + ref.startColumn })
      }
      const cachedFilteredOut = filterColumns.length ? hiddenRows.filter((r) => r > ref.startRow && r <= ref.endRow) : []
      for (const r of cachedFilteredOut) {
        delete rowData[r].hd
        if (!Object.keys(rowData[r]).length) delete rowData[r]
      }
      resources.filter[id] = { ref, filterColumns, cachedFilteredOut }
    }

    // 超级表（创建工作簿后再用 Facade 添加）
    for (const t of Object.values(ws.tables || {})) {
      const model = t.table || t
      const ref = model.tableRef || model.ref
      if (!ref) continue
      const range = a1ToRange(ref.includes(':') ? ref : `${ref}:${ref}`)
      post.tables.push({ sheetId: id, name: model.displayName || model.name, range, showHeader: model.headerRow !== false })
      if (model.totalsRow) warnings.push(`${ws.name}：超级表「${model.name}」的汇总行按普通单元格导入`)
    }

    // 批注
    if (Object.keys(notes).length) resources.note[id] = notes

    // 新式评论
    const threaded = extra.threadedComments || []
    if (threaded.length) {
      const roots = threaded.filter((c) => !c.parentId)
      resources.comment[id] = roots.map((root) => {
        const toComment = (c) => ({
          id: c.id,
          threadId: root.id,
          ref: root.ref,
          dT: toUniverTime(c.dT),
          personId: c.personId || 'excel-user',
          authorName: c.author || undefined,
          text: commentBody(c.text),
          attachments: [],
          unitId,
          subUnitId: id,
          ...(c.parentId ? { parentId: root.id } : {}),
        })
        return {
          ...toComment(root),
          resolved: root.done || undefined,
          children: threaded.filter((c) => c.parentId === root.id).map(toComment),
        }
      })
    }

    // 图片
    const images = ws.getImages?.() || []
    if (images.length) {
      const metrics = createSheetMetrics(sheet)
      const data = {}
      const order = []
      for (const img of images) {
        const media = wb.getImage(Number(img.imageId))
        if (!media?.buffer) continue
        const ext = (media.extension || 'png').replace('jpg', 'jpeg')
        const source = `data:image/${ext};base64,${bytesToBase64(new Uint8Array(media.buffer))}`
        const tl = img.range?.tl || {}
        const col = tl.nativeCol ?? Math.floor(tl.col ?? 0)
        const row = tl.nativeRow ?? Math.floor(tl.row ?? 0)
        const colOff = (tl.nativeColOff ?? 0) / EMU_PER_PX
        const rowOff = (tl.nativeRowOff ?? 0) / EMU_PER_PX
        const left = metrics.colLeft(col) + colOff
        const top = metrics.rowTop(row) + rowOff
        let width
        let height
        if (img.range?.ext?.width) {
          width = img.range.ext.width
          height = img.range.ext.height
        } else if (img.range?.br) {
          const br = img.range.br
          width = metrics.colLeft(br.nativeCol ?? 0) + (br.nativeColOff ?? 0) / EMU_PER_PX - left
          height = metrics.rowTop(br.nativeRow ?? 0) + (br.nativeRowOff ?? 0) / EMU_PER_PX - top
        }
        width = Math.max(1, Math.round(width || 100))
        height = Math.max(1, Math.round(height || 100))
        const toX = metrics.locateX(left + width)
        const toY = metrics.locateY(top + height)
        const from = { column: col, columnOffset: colOff, row, rowOffset: rowOff }
        const to = { column: toX.col, columnOffset: toX.offset, row: toY.row, rowOffset: toY.offset }
        const base = { flipY: false, flipX: false, angle: 0, skewX: 0, skewY: 0 }
        const drawingId = randomId(6)
        data[drawingId] = {
          unitId,
          subUnitId: id,
          drawingId,
          drawingType: 0,
          imageSourceType: 'URL',
          source,
          transform: { left, top, width, height, ...base },
          sheetTransform: { from, to, ...base },
          axisAlignSheetTransform: { from, to, ...base },
        }
        order.push(drawingId)
      }
      if (order.length) resources.drawing[id] = { data, order }
    }

    // 保护
    // Excel：保护工作表后，除了“未锁定”的单元格（含空单元格），其余都不能改。
    // - 没有未锁定的单元格 -> Univer 工作表保护
    // - 有未锁定的单元格 -> Univer 没有“保护整表但留几个洞”的能力，改为保护“未锁定区域以外的全部区域”
    if (ws.sheetProtection?.sheet) {
      const maxR = sheet.rowCount - 1
      const maxC = sheet.columnCount - 1
      const unlockedRects = cellsToRanges(unlockedCells)
      // 整列 / 整行设置为未锁定（空单元格也继承这个设置）
      ;(ws.columns || []).forEach((col, i) => {
        if (col?.protection?.locked === false) unlockedRects.push({ startRow: 0, endRow: maxR, startColumn: i, endColumn: i })
      })
      for (let r = 1; r <= ws.rowCount; r++) {
        const row = ws.findRow(r)
        if (row?.protection?.locked === false) unlockedRects.push({ startRow: r - 1, endRow: r - 1, startColumn: 0, endColumn: maxC })
      }
      if (!unlockedRects.length) {
        post.protectedSheets.push({ sheetId: id, allow: ws.sheetProtection })
      } else {
        const inUnlocked = (c) => unlockedRects.some((u) => c.row >= u.startRow && c.row <= u.endRow && c.col >= u.startColumn && c.col <= u.endColumn)
        // 未锁定的整行/整列里，单独设为“锁定”的单元格仍要保护
        const explicitLocked = cellsToRanges(lockedCells.filter(inUnlocked))
        const ranges = [...complementRects(unlockedRects, maxR, maxC), ...explicitLocked]
        if (ranges.length <= 60) {
          post.rangeProtections.push({ sheetId: id, ranges, name: `${ws.name} 保护区域` })
        } else {
          post.protectedSheets.push({ sheetId: id, allow: ws.sheetProtection })
          warnings.push(`${ws.name}：未锁定的单元格太零散，按整表保护导入`)
        }
      }
    }
  }

  // 定义名称（跳过 Excel 内置的 _xlnm.* 名称，如打印区域、筛选区域）
  for (const d of extras.definedNames) {
    if (!d.name || d.name.startsWith('_xlnm.')) continue
    const dnId = randomId(10)
    let formula = fromExcelFormula(d.formula)
    // 纯引用（Sheet1!$A$1:$B$2）不带等号，和 Univer 自己保存的格式一致
    if (/^=(?:'(?:[^']|'')+'|[^!=(),\s]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/i.test(formula)) formula = formula.slice(1)
    resources.definedName[dnId] = {
      id: dnId,
      name: d.name,
      formulaOrRefString: formula,
      localSheetId: d.localSheetId !== undefined ? (sheetIds[d.localSheetId] || 'AllDefaultWorkbook') : 'AllDefaultWorkbook',
      ...(d.hidden ? { hidden: true } : {}),
      ...(d.comment ? { comment: d.comment } : {}),
    }
  }

  if (!worksheets.length) throw new Error('这个文件里没有工作表')
  if (worksheets.every((_, i) => sheets[sheetIds[i]].hidden)) sheets[sheetIds[0]].hidden = 0

  const res = []
  const add = (name, value) => { if (Object.keys(value).length) res.push({ name, data: JSON.stringify(value) }) }
  add('SHEET_CONDITIONAL_FORMATTING_PLUGIN', resources.cf)
  add('SHEET_DATA_VALIDATION_PLUGIN', resources.dv)
  add('SHEET_FILTER_PLUGIN', resources.filter)
  add('SHEET_NOTE_PLUGIN', resources.note)
  add('SHEET_UNIVER_THREAD_COMMENT_PLUGIN', resources.comment)
  add('SHEET_DRAWING_PLUGIN', resources.drawing)
  add('SHEET_DEFINED_NAME_PLUGIN', resources.definedName)

  return {
    data: { id: unitId, name, sheetOrder: sheetIds, sheets, styles, resources: res },
    post,
    warnings,
  }
}

// 创建工作簿后才能做的事：超级表、工作表保护
async function applyPostSteps(univerAPI, fWorkbook, post, warnings) {
  for (const t of post.tables) {
    try {
      const ok = await fWorkbook.addTable(t.sheetId, t.name, t.range, undefined, { showHeader: t.showHeader })
      if (!ok) warnings.push(`超级表「${t.name}」没能创建（可能和筛选区域重叠）`)
    } catch (err) {
      warnings.push(`超级表「${t.name}」没能创建：${err.message}`)
    }
  }
  // 保护：创建后按 Excel 规则关掉“可编辑”权限（否则 Univer 默认允许创建者编辑）
  await withProtectionHookSuppressed(async () => {
    for (const { sheetId, ranges, name } of post.rangeProtections) {
      const fSheet = fWorkbook.getSheetBySheetId(sheetId)
      try {
        const fRanges = ranges.map((rg) => fSheet.getRange(rg.startRow, rg.startColumn, rg.endRow - rg.startRow + 1, rg.endColumn - rg.startColumn + 1))
        const rules = await fSheet.getWorksheetPermission().protectRanges([{ ranges: fRanges, options: { name } }])
        for (const rule of rules || []) await applyExcelRangePoints(univerAPI, rule)
      } catch (err) {
        warnings.push(`区域保护没能设置：${err.message}`)
      }
    }
    for (const { sheetId, allow } of post.protectedSheets) {
      try {
        const fSheet = fWorkbook.getSheetBySheetId(sheetId)
        await fSheet.getWorksheetPermission().protect()
        await applyExcelWorksheetPoints(univerAPI, fSheet, allow)
      } catch (err) {
        warnings.push(`工作表保护没能设置：${err.message}`)
      }
    }
  })
}

export async function importXlsxFile(univerAPI, file) {
  const { data, post, warnings } = await xlsxToWorkbookData(await file.arrayBuffer(), file.name.replace(/\.xlsx$/i, ''))
  const current = univerAPI.getActiveWorkbook()
  if (current) univerAPI.disposeUnit(current.getId())
  const workbook = univerAPI.createWorkbook(data)
  await applyPostSteps(univerAPI, workbook, post, warnings)
  // 重新计算所有公式（动态数组的溢出结果需要计算后才会出现）
  try {
    const formula = univerAPI.getFormula()
    formula.executeCalculation()
    await formula.onCalculationResultApplied(10000)
  } catch { /* 计算超时不影响导入结果 */ }
  return { workbook, warnings }
}

// 弹出文件选择框，读取 .xlsx 并替换当前工作簿
export function pickAndImportXlsx(univerAPI) {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    input.style.display = 'none'
    input.addEventListener('change', async () => {
      const file = input.files?.[0]
      input.remove()
      if (!file) return resolve(null)
      try {
        resolve(await importXlsxFile(univerAPI, file))
      } catch (err) {
        reject(err)
      }
    })
    document.body.appendChild(input)
    input.click()
  })
}
