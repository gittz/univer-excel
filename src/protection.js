// 让 Univer 的保护按 Excel 的规则生效
//
// Univer 开源版的保护默认是“只有创建者能编辑，其他人只读”，当前用户就是创建者，所以受保护的单元格照样能改。
// Excel 的规则是“一旦保护，谁都不能改，直到撤销保护”。这里在保护创建后把“可编辑”等权限点关掉，
// 只保留查看、复制、选择和撤销保护。

let suppress = 0 // 导入时由导入流程自己设置权限点，暂停自动处理

export function withProtectionHookSuppressed(fn) {
  suppress++
  return Promise.resolve().then(fn).finally(() => { suppress-- })
}

/**
 * 按 Excel 的工作表保护选项设置 Univer 工作表权限点
 * @param {object} allow ExcelJS 的 sheetProtection 模型：formatCells / insertRows / sort / autoFilter 等为 true 表示允许
 */
export async function applyExcelWorksheetPoints(univerAPI, fSheet, allow = {}) {
  const P = univerAPI.Enum.WorksheetPermissionPoint
  const perm = fSheet.getWorksheetPermission()
  const points = [
    [P.Edit, false],
    [P.SetCellValue, false],
    [P.SetCellStyle, !!allow.formatCells],
    [P.SetRowStyle, !!allow.formatRows],
    [P.SetColumnStyle, !!allow.formatColumns],
    [P.InsertRow, !!allow.insertRows],
    [P.InsertColumn, !!allow.insertColumns],
    [P.DeleteRow, !!allow.deleteRows],
    [P.DeleteColumn, !!allow.deleteColumns],
    [P.Sort, !!allow.sort],
    [P.Filter, !!allow.autoFilter],
    [P.PivotTable, !!allow.pivotTables],
    [P.InsertHyperlink, !!allow.insertHyperlinks],
    [P.EditExtraObject, allow.objects === false],
    [P.SelectProtectedCells, allow.selectLockedCells !== false],
    [P.SelectUnProtectedCells, allow.selectUnlockedCells !== false],
    [P.View, true],
    [P.Copy, true],
    [P.DeleteProtection, true],
  ]
  for (const [point, value] of points) {
    if (point) await perm.setPoint(point, value)
  }
}

export async function applyExcelRangePoints(univerAPI, rule) {
  const P = univerAPI.Enum.RangePermissionPoint
  await rule.setPoint(P.Edit, false)
  await rule.setPoint(P.View, true)
  await rule.setPoint(P.Delete, true)
}

// 对工作簿里所有“还按 Univer 默认规则生效”的保护补上 Excel 规则
export async function enforceExcelProtection(univerAPI, fWorkbook = univerAPI.getActiveWorkbook()) {
  if (!fWorkbook) return
  const WP = univerAPI.Enum.WorksheetPermissionPoint
  const RP = univerAPI.Enum.RangePermissionPoint
  for (const fSheet of fWorkbook.getSheets()) {
    const perm = fSheet.getWorksheetPermission()
    if (perm.isProtected() && perm.getPoint(WP.Edit)) await applyExcelWorksheetPoints(univerAPI, fSheet)
    let rules = []
    try { rules = await perm.listRangeProtectionRules() } catch { /* ignore */ }
    for (const rule of rules) {
      if (rule.getPoint(RP.Edit)) await applyExcelRangePoints(univerAPI, rule)
    }
  }
}

// 界面上新建保护时也按 Excel 规则生效
export function installExcelProtection(univerAPI) {
  // 界面菜单走 command，Facade 的 protect() 直接走 mutation，两种都要听
  const ids = new Set([
    'sheet.command.add-range-protection', 'sheet.command.add-worksheet-protection',
    'sheet.command.set-worksheet-protection', 'sheet.command.set-range-protection',
    'sheet.mutation.add-range-protection', 'sheet.mutation.add-worksheet-protection',
    'sheet.mutation.set-range-protection', 'sheet.mutation.set-worksheet-protection',
  ])
  return univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
    if (!ids.has(event.id) || suppress) return
    setTimeout(() => {
      if (!suppress) enforceExcelProtection(univerAPI).catch((err) => console.warn('[protection]', err))
    }, 0)
  })
}
