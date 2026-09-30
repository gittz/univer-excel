// 纯前端把 Univer 工作簿导出为 .xlsx（只用开源的 ExcelJS / JSZip，不依赖 Univer Pro 或服务端）
import ExcelJS from 'exceljs'
import {
  applyUniverStyleToExcelCell, univerFontToExcel, createSheetMetrics, parseResources, docToText,
  rangeToA1, cellRef, complementRects, pxToColWidth, pxToPt, toArgb, EMU_PER_PX, bytesToBase64,
} from './xlsx/utils.js'
import { toExcelFormula } from './xlsx/formula.js'
import { univerCfToExcel } from './xlsx/conditional-format.js'
import { univerDvToExcel } from './xlsx/data-validation.js'
import { patchXlsxExport } from './xlsx/xml-parts.js'

const R = {
  CF: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
  DV: 'SHEET_DATA_VALIDATION_PLUGIN',
  FILTER: 'SHEET_FILTER_PLUGIN',
  TABLE: 'SHEET_TABLE_PLUGIN',
  NOTE: 'SHEET_NOTE_PLUGIN',
  COMMENT: 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN',
  DRAWING: 'SHEET_DRAWING_PLUGIN',
  DEFINED_NAME: 'SHEET_DEFINED_NAME_PLUGIN',
}

function scalarValue(cellData) {
  const v = cellData.v
  if (v === undefined || v === null) return null
  if (cellData.t === 3) return v === 1 || v === true || v === 'TRUE' || v === '1'
  if (cellData.t === 4 || cellData.t === 1) return String(v)
  return v
}

// 单元格文档（p）-> ExcelJS 值：超链接 / 富文本 / 纯文本
function docValue(p, sheetNames) {
  const body = p?.body
  const text = docToText(p)
  const link = (body?.customRanges || []).find((r) => r.rangeType === 0 && r.properties?.url)
  if (link) {
    let url = link.properties.url
    // Univer 内部链接 '#gid=sheetId&range=A1' -> Excel '#'Sheet'!A1'
    if (url.startsWith('#')) {
      const params = new URLSearchParams(url.slice(1))
      const sheetName = sheetNames[params.get('gid')]
      const range = params.get('range') || 'A1'
      if (sheetName) url = `#'${sheetName.replace(/'/g, "''")}'!${range}`
      else if (params.get('rangeid') || params.get('range')) url = `#${range}`
    }
    return { text, hyperlink: url, tooltip: link.properties.tooltip || undefined }
  }
  const runs = body?.textRuns || []
  if (runs.length) {
    const stream = body.dataStream.replace(/\r\n$/, '')
    const pieces = []
    let pos = 0
    const sorted = [...runs].sort((a, b) => a.st - b.st)
    for (const run of sorted) {
      if (run.st > pos) pieces.push({ text: stream.slice(pos, run.st) })
      const t = stream.slice(Math.max(run.st, pos), run.ed)
      if (t) pieces.push({ text: t, font: univerFontToExcel(run.ts) })
      pos = Math.max(pos, run.ed)
    }
    if (pos < stream.length) pieces.push({ text: stream.slice(pos) })
    const richText = pieces
      .map((x) => ({ ...x, text: x.text.replace(/[\b]/g, '').replace(/\r/g, '\n') }))
      .filter((x) => x.text)
    if (richText.length) return { richText }
  }
  return text || null
}

async function imageSourceToBase64(source) {
  if (!source) return null
  if (source.startsWith('data:')) {
    const m = source.match(/^data:image\/(png|jpe?g|gif);base64,(.*)$/i)
    if (!m) return null
    return { base64: m[2], extension: m[1].toLowerCase() === 'jpg' ? 'jpeg' : m[1].toLowerCase() }
  }
  // 普通 URL：尽量下载下来（跨域不允许时会失败）
  const res = await fetch(source)
  const blob = await res.blob()
  const ext = (blob.type.split('/')[1] || 'png').replace('jpg', 'jpeg')
  if (!['png', 'jpeg', 'gif'].includes(ext)) return null
  return { base64: bytesToBase64(new Uint8Array(await blob.arrayBuffer())), extension: ext }
}

const toTime = (dT) => {
  if (!dT) return ''
  return String(dT).replace('T', ' ').slice(0, 16)
}

export async function workbookToXlsxBuffer(univerAPI, fWorkbook) {
  const snapshot = fWorkbook.save()
  const res = parseResources(snapshot)
  const styles = snapshot.styles || {}
  const warnings = []
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Univer'

  const sheetNames = Object.fromEntries(snapshot.sheetOrder.map((id) => [id, snapshot.sheets[id]?.name || id]))
  const sheetIndex = Object.fromEntries(snapshot.sheetOrder.map((id, i) => [id, i]))
  const patch = { definedNames: [], filters: {}, validations: {}, unlockAllColumns: [] }
  let currentUser = null
  try { currentUser = univerAPI.getUserManager().getCurrentUser() } catch { /* ignore */ }
  const authorOf = (c) => c.authorName || (currentUser && c.personId === currentUser.userID ? currentUser.name : '') || c.personId || ''

  for (const [index, sheetId] of snapshot.sheetOrder.entries()) {
    const sheet = snapshot.sheets[sheetId]
    if (!sheet) continue
    const name = sheet.name || sheetId
    const fSheet = fWorkbook.getSheetBySheetId(sheetId)
    const metrics = createSheetMetrics(sheet)
    const cellData = sheet.cellData || {}

    const ws = wb.addWorksheet(name, {
      state: sheet.hidden === 1 ? 'hidden' : 'visible',
      properties: { defaultRowHeight: pxToPt(sheet.defaultRowHeight || 24) },
    })
    const tab = toArgb(sheet.tabColor)
    if (tab) ws.properties.tabColor = { argb: tab }
    const view = {}
    const f = sheet.freeze
    if (f && (f.xSplit > 0 || f.ySplit > 0)) Object.assign(view, { state: 'frozen', xSplit: f.xSplit || 0, ySplit: f.ySplit || 0 })
    if (sheet.showGridlines === 0) view.showGridLines = false
    if (sheet.zoomRatio && sheet.zoomRatio !== 1) view.zoomScale = Math.round(sheet.zoomRatio * 100)
    if (Object.keys(view).length) ws.views = [view]

    // 行高列宽、隐藏
    ws.properties.defaultColWidth = pxToColWidth(sheet.defaultColumnWidth || 88)
    for (const [c, col] of Object.entries(sheet.columnData || {})) {
      const excelCol = ws.getColumn(Number(c) + 1)
      if (col.w) excelCol.width = pxToColWidth(col.w)
      if (col.hd === 1) excelCol.hidden = true
    }
    for (const [r, row] of Object.entries(sheet.rowData || {})) {
      const excelRow = ws.getRow(Number(r) + 1)
      const h = row.ah && row.ia !== 0 ? row.ah : row.h
      if (h) excelRow.height = pxToPt(h)
      if (row.hd === 1) excelRow.hidden = true
    }

    // 单元格
    let maxRow = 0
    let maxCol = 0
    const cellImages = []
    const spillCovered = new Set()
    for (const [r, row] of Object.entries(cellData)) {
      for (const [c, cd] of Object.entries(row || {})) {
        if (!cd) continue
        const rowIdx = Number(r)
        const colIdx = Number(c)
        maxRow = Math.max(maxRow, rowIdx)
        maxCol = Math.max(maxCol, colIdx)
        const cell = ws.getCell(rowIdx + 1, colIdx + 1)

        let formula = cd.f
        if (!formula && cd.si && fSheet) formula = fSheet.getRange(rowIdx, colIdx).getFormula() // 拖拽填充产生的共享公式
        let value = cd.p ? docValue(cd.p, sheetNames) : scalarValue(cd)
        if (cd.p && value === null) value = scalarValue(cd)

        if (formula) {
          const result = scalarValue(cd)
          const fv = { formula: toExcelFormula(formula), result: result ?? undefined }
          // 动态数组：结果溢出到了右边 / 下边的空单元格
          if (fSheet) {
            const isFree = (rr, cc) => !cellData[rr]?.[cc]
            let w = 1
            while (w < 500 && isFree(rowIdx, colIdx + w) && fSheet.getRange(rowIdx, colIdx + w).getValue() !== null && fSheet.getRange(rowIdx, colIdx + w).getValue() !== undefined) w++
            let h = 1
            while (h < 5000 && isFree(rowIdx + h, colIdx) && fSheet.getRange(rowIdx + h, colIdx).getValue() !== null && fSheet.getRange(rowIdx + h, colIdx).getValue() !== undefined) h++
            if (w > 1 || h > 1) {
              const values = fSheet.getRange(rowIdx, colIdx, h, w).getValues()
              fv.shareType = 'array'
              fv.ref = rangeToA1({ startRow: rowIdx, startColumn: colIdx, endRow: rowIdx + h - 1, endColumn: colIdx + w - 1 })
              fv.result = values[0][0] ?? undefined
              for (let i = 0; i < h; i++) {
                for (let j = 0; j < w; j++) {
                  if (!i && !j) continue
                  spillCovered.add(`${rowIdx + i},${colIdx + j}`)
                  ws.getCell(rowIdx + i + 1, colIdx + j + 1).value = values[i][j]
                }
              }
              maxRow = Math.max(maxRow, rowIdx + h - 1)
              maxCol = Math.max(maxCol, colIdx + w - 1)
            }
          }
          cell.value = fv
        } else if (value !== null && value !== '') {
          cell.value = value
        }

        const style = typeof cd.s === 'string' ? styles[cd.s] : cd.s
        applyUniverStyleToExcelCell(cell, style)

        // 单元格内图片
        const drawings = cd.p?.drawings
        if (drawings) {
          for (const id of cd.p.drawingsOrder || Object.keys(drawings)) {
            const d = drawings[id]
            if (d?.source) cellImages.push({ row: rowIdx, col: colIdx, source: d.source, width: d.transform?.width || d.docTransform?.size?.width || 64, height: d.transform?.height || d.docTransform?.size?.height || 40 })
          }
        }
      }
    }
    if (spillCovered.size) warnings.push(`${name}：动态数组公式导出为数组公式（Excel 里显示为 {=...}，结果一致）`)

    // 合并单元格
    for (const m of sheet.mergeData || []) {
      ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1)
    }
    const limitRow = Math.max(maxRow + 1, (sheet.rowCount || 1000) - 1)
    const limitCol = Math.max(maxCol + 1, (sheet.columnCount || 20) - 1)

    // 条件格式
    for (const cf of univerCfToExcel(res[R.CF]?.[sheetId] || [], { sheetName: name, warnings })) ws.addConditionalFormatting(cf)

    // 数据验证
    const dvList = univerDvToExcel(res[R.DV]?.[sheetId] || [], { sheetName: name, maxRow: limitRow, maxCol: limitCol, warnings })
    if (dvList.length) patch.validations[index] = dvList

    // 筛选
    const filter = res[R.FILTER]?.[sheetId]
    if (filter?.ref) {
      ws.autoFilter = rangeToA1(filter.ref)
      const columns = (filter.filterColumns || []).map((c) => ({ ...c, colId: c.colId - filter.ref.startColumn })).filter((c) => c.colId >= 0)
      if ((filter.filterColumns || []).some((c) => c.colorFilters)) warnings.push(`${name}：按颜色筛选的条件暂不支持导出`)
      if (columns.length) patch.filters[index] = { ref: ws.autoFilter, columns }
      for (const r of filter.cachedFilteredOut || []) ws.getRow(r + 1).hidden = true
    }

    // 超级表
    for (const t of res[R.TABLE]?.[sheetId]?.tables || []) {
      const rg = t.range
      const showHeader = t.options?.showHeader !== false
      const showFooter = !!t.options?.showFooter
      const firstDataRow = rg.startRow + (showHeader ? 1 : 0)
      const lastDataRow = rg.endRow - (showFooter ? 1 : 0)
      const rows = []
      for (let r = firstDataRow; r <= lastDataRow; r++) {
        const row = []
        for (let c = rg.startColumn; c <= rg.endColumn; c++) row.push(ws.getCell(r + 1, c + 1).value)
        rows.push(row)
      }
      const tableName = String(t.name || `Table${index + 1}`).replace(/[^\p{L}\p{N}_.]/gu, '_').replace(/^(\d)/, '_$1')
      const columns = (t.columns || []).map((col, i) => ({
        name: col.displayName || ws.getCell(rg.startRow + 1, rg.startColumn + i + 1).text || `列${i + 1}`,
        filterButton: t.options?.showAutoFilter !== false && col.showFilterButton !== false,
      }))
      ws.addTable({
        name: tableName,
        displayName: tableName,
        ref: cellRef(rg.startRow, rg.startColumn),
        headerRow: showHeader,
        totalsRow: false,
        style: { theme: 'TableStyleMedium2', showRowStripes: true },
        columns,
        rows: rows.length ? rows : [columns.map(() => null)],
      })
      if (showFooter) warnings.push(`${name}：超级表「${t.name}」的汇总行按普通单元格导出`)
    }

    // 批注（Excel 的 note）
    const notes = res[R.NOTE]?.[sheetId] || {}
    const noteText = {}
    for (const [r, cols] of Object.entries(notes)) {
      for (const [c, note] of Object.entries(cols || {})) {
        if (note?.note) noteText[`${r},${c}`] = note.note
      }
    }
    // 评论：Excel 的新式评论 ExcelJS 写不了，按“作者 时间：内容”合并到批注里
    const comments = res[R.COMMENT]?.[sheetId] || []
    if (comments.length) warnings.push(`${name}：评论导出为 Excel 批注（含作者、时间和回复）`)
    for (const cm of comments) {
      let rc
      try {
        const m = String(cm.ref).match(/^([A-Z]+)(\d+)$/i)
        let col = 0
        for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
        rc = `${Number(m[2]) - 1},${col - 1}`
      } catch { continue }
      const lines = [cm, ...(cm.children || [])].map((x) => {
        const who = authorOf(x)
        const body = String(x.text?.dataStream || '').replace(/\r\n$/, '').replace(/\r/g, '\n')
        return `${who}${who ? ' ' : ''}${toTime(x.dT)}：${body}`
      })
      const text = (cm.resolved ? '[已解决] ' : '') + lines.join('\n')
      noteText[rc] = noteText[rc] ? `${noteText[rc]}\n---\n${text}` : text
    }
    for (const [rc, text] of Object.entries(noteText)) {
      const [r, c] = rc.split(',').map(Number)
      ws.getCell(r + 1, c + 1).note = text
    }

    // 图片：浮动图片 + 单元格内图片
    const drawingData = res[R.DRAWING]?.[sheetId]
    const floatImages = drawingData ? (drawingData.order || Object.keys(drawingData.data || {})).map((id) => drawingData.data?.[id]).filter(Boolean) : []
    for (const d of floatImages) {
      if (d.drawingType !== undefined && d.drawingType !== 0) { warnings.push(`${name}：非图片类型的绘图对象暂不支持导出`); continue }
      try {
        const img = await imageSourceToBase64(d.source)
        if (!img) { warnings.push(`${name}：有一张图片格式不支持（仅支持 png/jpeg/gif）`); continue }
        const from = d.sheetTransform?.from || { row: d.row ?? 0, rowOffset: d.rowOffset ?? 0, column: d.column ?? 0, columnOffset: d.columnOffset ?? 0 }
        const width = d.transform?.width ?? d.width ?? 100
        const height = d.transform?.height ?? d.height ?? 100
        const imageId = wb.addImage(img)
        ws.addImage(imageId, {
          tl: { nativeCol: from.column, nativeColOff: Math.round(from.columnOffset * EMU_PER_PX), nativeRow: from.row, nativeRowOff: Math.round(from.rowOffset * EMU_PER_PX) },
          ext: { width, height },
          editAs: 'oneCell',
        })
      } catch (err) {
        warnings.push(`${name}：一张图片无法下载（${err.message}），已跳过`)
      }
    }
    for (const ci of cellImages) {
      try {
        const img = await imageSourceToBase64(ci.source)
        if (!img) continue
        const cw = metrics.colW(ci.col)
        const rh = metrics.rowH(ci.row)
        const scale = Math.min(1, cw / ci.width, rh / ci.height)
        const imageId = wb.addImage(img)
        ws.addImage(imageId, {
          tl: { nativeCol: ci.col, nativeColOff: 0, nativeRow: ci.row, nativeRowOff: 0 },
          ext: { width: Math.max(1, Math.floor(ci.width * scale)), height: Math.max(1, Math.floor(ci.height * scale)) },
          editAs: 'oneCell',
        })
      } catch (err) {
        warnings.push(`${name}：一张单元格图片无法下载（${err.message}），已跳过`)
      }
    }
    if (cellImages.length) warnings.push(`${name}：单元格内图片导出为贴在单元格上的浮动图片`)

    // 保护
    if (fSheet) {
      const perm = fSheet.getWorksheetPermission()
      const sheetProtected = perm.isProtected?.()
      let rangeRules = []
      try { rangeRules = await perm.listRangeProtectionRules() } catch { /* ignore */ }
      if (sheetProtected) {
        // 把 Univer 工作表权限点换成 Excel 的保护选项（true 表示允许）
        const P = univerAPI.Enum.WorksheetPermissionPoint
        const can = (point) => { try { return !!perm.getPoint(point) } catch { return false } }
        await ws.protect('', {
          selectLockedCells: can(P.SelectProtectedCells),
          selectUnlockedCells: can(P.SelectUnProtectedCells),
          formatCells: can(P.SetCellStyle),
          formatRows: can(P.SetRowStyle),
          formatColumns: can(P.SetColumnStyle),
          insertRows: can(P.InsertRow),
          insertColumns: can(P.InsertColumn),
          deleteRows: can(P.DeleteRow),
          deleteColumns: can(P.DeleteColumn),
          sort: can(P.Sort),
          autoFilter: can(P.Filter),
          insertHyperlinks: can(P.InsertHyperlink),
        })
      } else if (rangeRules.length) {
        // Excel 没有“只保护部分区域”的开关：保护整张表，受保护区域以外的单元格（包括空单元格）设为“未锁定”。
        // Excel 默认所有单元格都是锁定的；ExcelJS 只会写出有单元格的行，所以整行的样式没法落到空行上。
        // 两种写法取需要写的单元格更少的一种：
        //   A. 所有列设为未锁定（列样式覆盖到空单元格），再把受保护区域的单元格逐个设为锁定
        //   B. 保持默认锁定，只把未锁定区域的单元格逐个设为未锁定
        const clip = (rg) => ({ startRow: rg.startRow, endRow: Math.min(rg.endRow, limitRow), startColumn: rg.startColumn, endColumn: Math.min(rg.endColumn, limitCol) })
        const protectedRects = rangeRules.flatMap((rule) => rule.ranges.map((fr) => clip(fr.getRange())))
        const unlockedRects = complementRects(protectedRects, limitRow, limitCol)
        const area = (rects) => rects.reduce((sum, r) => sum + (r.endRow - r.startRow + 1) * (r.endColumn - r.startColumn + 1), 0)
        // ExcelJS 里新建的单元格会和列共用同一个样式对象，直接改 protection 会互相影响，所以换成新对象
        // 锁定就是 Excel 的默认值：去掉 protection 让它回到默认（ExcelJS 写 locked=true 时只写 applyProtection、不写 locked="1"，有的软件会读错）
        const setLocked = (cell, locked) => {
          const { protection, ...rest } = cell.style || {}
          cell.style = locked ? rest : { ...rest, protection: { locked: false } }
        }
        const inRects = (rects, r, c) => rects.some((x) => r >= x.startRow && r <= x.endRow && c >= x.startColumn && c <= x.endColumn)
        if (area(protectedRects) <= area(unlockedRects)) {
          // 整列设为未锁定放到 XML 修补阶段做（ExcelJS 设置列样式会顺带把列宽改成 9）
          patch.unlockAllColumns.push(index)
          // 已有内容的单元格有自己的样式，不继承列样式，要单独设置
          for (let r = 0; r <= maxRow; r++) {
            for (let c = 0; c <= maxCol; c++) {
              if (!inRects(protectedRects, r, c)) setLocked(ws.getCell(r + 1, c + 1), false)
            }
          }
          for (const rg of protectedRects) {
            for (let r = rg.startRow; r <= rg.endRow; r++) {
              for (let c = rg.startColumn; c <= rg.endColumn; c++) setLocked(ws.getCell(r + 1, c + 1), true)
            }
          }
        } else {
          for (const rg of unlockedRects) {
            for (let r = rg.startRow; r <= rg.endRow; r++) {
              for (let c = rg.startColumn; c <= rg.endColumn; c++) setLocked(ws.getCell(r + 1, c + 1), false)
            }
          }
        }
        await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true })
        warnings.push(`${name}：区域保护导出为“工作表保护 + 其余单元格解除锁定”`)
      }
    }
  }

  // 定义名称
  for (const d of Object.values(res[R.DEFINED_NAME] || {})) {
    if (!d?.name) continue
    const local = d.localSheetId && d.localSheetId !== 'AllDefaultWorkbook' ? sheetIndex[d.localSheetId] : undefined
    patch.definedNames.push({
      name: d.name,
      formula: toExcelFormula(String(d.formulaOrRefString || '').replace(/\[[^\]]*\]/g, '')),
      localSheetId: local,
      hidden: !!d.hidden,
      comment: d.comment || undefined,
    })
  }

  const raw = await wb.xlsx.writeBuffer()
  const buffer = await patchXlsxExport(raw, patch)
  return { buffer, warnings, name: snapshot.name || 'workbook' }
}

export async function exportWorkbookToXlsx(univerAPI, fWorkbook) {
  const { buffer, warnings, name } = await workbookToXlsxBuffer(univerAPI, fWorkbook)
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const fileName = `${name}.xlsx`
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
  return { fileName, warnings }
}
