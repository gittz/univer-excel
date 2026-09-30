// Univer 开源版全功能示例（Presets 方式，不使用任何 @univerjs-pro 包）
// 文档：https://docs.univer.ai/guides/sheets/getting-started/quickstart
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'

import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import sheetsCFZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import sheetsDVZhCN from '@univerjs/preset-sheets-data-validation/locales/zh-CN'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import sheetsFilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import sheetsSortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import sheetsTableZhCN from '@univerjs/preset-sheets-table/locales/zh-CN'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import sheetsNoteZhCN from '@univerjs/preset-sheets-note/locales/zh-CN'
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment'
import sheetsThreadCommentZhCN from '@univerjs/preset-sheets-thread-comment/locales/zh-CN'
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import sheetsHyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import sheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import sheetsFindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN'
import { UniverWatermarkPlugin } from '@univerjs/watermark'
import '@univerjs/watermark/facade'

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'
import '@univerjs/preset-sheets-sort/lib/index.css'
import '@univerjs/preset-sheets-table/lib/index.css'
import '@univerjs/preset-sheets-note/lib/index.css'
import '@univerjs/preset-sheets-thread-comment/lib/index.css'
import '@univerjs/preset-sheets-hyper-link/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-find-replace/lib/index.css'

import { exportWorkbookToXlsx } from './export-xlsx.js'
import { pickAndImportXlsx } from './import-xlsx.js'
import { createDemoWorkbook } from './demo-data.js'
import { installExcelProtection, enforceExcelProtection } from './protection.js'

const { univerAPI } = createUniver({
  locale: LocaleType.ZH_CN,
  locales: {
    [LocaleType.ZH_CN]: mergeLocales(
      sheetsCoreZhCN,
      sheetsCFZhCN,
      sheetsDVZhCN,
      sheetsFilterZhCN,
      sheetsSortZhCN,
      sheetsTableZhCN,
      sheetsNoteZhCN,
      sheetsThreadCommentZhCN,
      sheetsHyperLinkZhCN,
      sheetsDrawingZhCN,
      sheetsFindReplaceZhCN,
    ),
  },
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
    UniverSheetsConditionalFormattingPreset(),
    UniverSheetsDataValidationPreset(),
    UniverSheetsFilterPreset(),
    UniverSheetsSortPreset(),
    UniverSheetsTablePreset(),
    UniverSheetsNotePreset(),
    UniverSheetsThreadCommentPreset(),
    UniverSheetsHyperLinkPreset(),
    UniverSheetsDrawingPreset(),
    UniverSheetsFindReplacePreset(),
  ],
  plugins: [UniverWatermarkPlugin],
})

// 保护按 Excel 规则生效：受保护的单元格谁都不能改（包括创建者），直到撤销保护
installExcelProtection(univerAPI)

createDemoWorkbook(univerAPI)
  .then(() => enforceExcelProtection(univerAPI))
  .catch((err) => console.error('[demo]', err))

// 工具栏「导入 Excel」：纯前端读取 .xlsx 并替换当前工作簿
univerAPI.createMenu({
  id: 'custom.import-xlsx',
  title: '导入 Excel',
  tooltip: '打开一个 .xlsx 文件（会替换当前工作簿）',
  order: -2,
  action: async () => {
    try {
      const result = await pickAndImportXlsx(univerAPI)
      if (!result) return
      const { workbook, warnings } = result
      univerAPI.showMessage({ content: `已导入：${workbook.getName()}`, type: 'success' })
      if (warnings.length) console.warn('[import-xlsx]', warnings)
    } catch (err) {
      console.error('[import-xlsx]', err)
      univerAPI.showMessage({ content: `导入失败：${err.message || err}`, type: 'error' })
    }
  },
}).appendTo('ribbon.start.history')

// 工具栏「导出 Excel」：纯前端导出 .xlsx
univerAPI.createMenu({
  id: 'custom.export-xlsx',
  title: '导出 Excel',
  tooltip: '把当前工作簿导出为 .xlsx 文件',
  order: -1,
  action: async () => {
    const fWorkbook = univerAPI.getActiveWorkbook()
    if (!fWorkbook) return
    try {
      const { fileName, warnings } = await exportWorkbookToXlsx(univerAPI, fWorkbook)
      univerAPI.showMessage({ content: `已导出：${fileName}`, type: 'success' })
      if (warnings.length) console.warn('[export-xlsx]', warnings)
    } catch (err) {
      console.error('[export-xlsx]', err)
      univerAPI.showMessage({ content: `导出失败：${err.message || err}`, type: 'error' })
    }
  },
}).appendTo('ribbon.start.history')

// 「数据」标签页里的“分列”：开源版有分列能力（FRange.splitTextToColumns），但没有菜单入口，这里补一个
univerAPI.createMenu({
  id: 'custom.split-text-to-columns',
  title: '分列',
  tooltip: '按分隔符把选中的一列文字拆成多列',
  action: () => {
    const range = univerAPI.getActiveWorkbook()?.getActiveRange()
    if (!range) return
    if (range.getWidth() !== 1) {
      univerAPI.showMessage({ content: '请先选中一列再分列', type: 'warning' })
      return
    }
    const input = window.prompt('分隔符（留空自动识别；可填 , ; 空格 tab 或任意字符）', '')
    if (input === null) return
    const d = input === '' ? null : input.toLowerCase()
    if (d === null) range.splitTextToColumns(false)
    else if (d === ',' || d === '，') range.splitTextToColumns(false, 2)
    else if (d === ';' || d === '；') range.splitTextToColumns(false, 4)
    else if (d === ' ' || d === '空格') range.splitTextToColumns(false, 8)
    else if (d === 'tab' || d === '\t') range.splitTextToColumns(false, 1)
    else range.splitTextToColumns(false, 16, input)
  },
}).appendTo('ribbon.data.others')

// 「视图」标签页里的水印开关（水印是显示层效果，Excel 没有对应概念，不会被导出）
let watermarkOn = false
univerAPI.createMenu({
  id: 'custom.toggle-watermark',
  title: '水印',
  tooltip: '显示 / 隐藏文字水印',
  action: () => {
    if (watermarkOn) {
      univerAPI.deleteWatermark()
      watermarkOn = false
      return
    }
    const text = window.prompt('水印文字', '内部资料 请勿外传')
    if (!text) return
    univerAPI.addWatermark('text', { content: text, fontSize: 16, color: 'rgb(0,0,0)', opacity: 0.12, repeat: true, spacingX: 160, spacingY: 120, rotate: -20 })
    watermarkOn = true
  },
}).appendTo('ribbon.view.others')

// 方便在浏览器控制台里调用 Facade API
window.univerAPI = univerAPI
