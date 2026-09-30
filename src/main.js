// Univer 开源版全功能示例（Presets 方式，不使用任何 @univerjs-pro 包）
// 文档：https://docs.univer.ai/guides/sheets/getting-started/quickstart
import { createUniver, LocaleType, mergeLocales, ICommandService, CommandType } from '@univerjs/presets'

import { UniverSheetsCorePreset, IMenuManagerService, MenuItemType, RibbonStartGroup } from '@univerjs/preset-sheets-core'
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
import { registerCustomIcons } from './ui-icons.js'

const { univer, univerAPI } = createUniver({
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
    // grid：和官方 demo 一样的 Excel 风格工具栏（标签靠左、两排按钮、常用功能大图标）
    UniverSheetsCorePreset({ container: 'app', ribbonType: 'grid' }),
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

registerCustomIcons(univer, univerAPI)

// 保护按 Excel 规则生效：受保护的单元格谁都不能改（包括创建者），直到撤销保护
installExcelProtection(univerAPI)

createDemoWorkbook(univerAPI)
  .then(() => enforceExcelProtection(univerAPI))
  .catch((err) => console.error('[demo]', err))

// 「开始」标签页的「文件 ▾」下拉：导入 / 导出 Excel（纯前端，不依赖服务端）
async function importExcel() {
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
}

async function exportExcel() {
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
}

// Facade 的 createSubmenu 在工具栏里弹不出选项，所以按 Univer 原生下拉按钮（如「冻结」）的写法注册：
// 菜单项类型 SUBITEMS + selections，选中某项时执行对应命令
const injector = univer.__getInjector()
const commandService = injector.get(ICommandService)
commandService.registerCommand({ id: 'custom.command.import-xlsx', type: CommandType.COMMAND, handler: () => { importExcel(); return true } })
commandService.registerCommand({ id: 'custom.command.export-xlsx', type: CommandType.COMMAND, handler: () => { exportExcel(); return true } })
injector.get(IMenuManagerService).mergeMenu({
  [RibbonStartGroup.OTHERS]: {
    'custom.file': {
      order: 1,
      // 在 grid 工具栏里占两行的大按钮、显示文字，紧挨「保护」（第 1 列）
      gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
      menuItemFactory: () => ({
        id: 'custom.file',
        type: MenuItemType.SUBITEMS,
        icon: 'CustomFileIcon',
        title: '文件',
        tooltip: '导入 / 导出 Excel',
        selections: [
          { id: 'custom.command.import-xlsx', value: 'import', label: '导入 Excel', icon: 'CustomImportIcon' },
          { id: 'custom.command.export-xlsx', value: 'export', label: '导出 Excel', icon: 'CustomExportIcon' },
        ],
      }),
    },
  },
})

// 「数据」标签页里的“分列”：开源版有分列能力（FRange.splitTextToColumns），但没有菜单入口，这里补一个
univerAPI.createMenu({
  id: 'custom.split-text-to-columns',
  title: '分列',
  tooltip: '按分隔符把选中的一列文字拆成多列',
  icon: 'CustomSplitColumnsIcon',
  gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
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
  icon: 'CustomWatermarkIcon',
  gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
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
