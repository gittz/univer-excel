// 示例工作簿：把开源版的主要功能都用上一遍，方便直接体验和测试导入导出
const CHART_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAAA4CAIAAACNJ2r1AAAAiUlEQVR42u3ZIRWAQBRE0TU0IgEVCIDAUIQuG4JYOASQYMVHfDjcdybBtVNONSsIAAECBAgQIECABAgQIECAAAECpFSgvo7hAQIECBAgQIAAAQIECNAfgIZ1Dw8QIECAAAECBAgQIEBvBzq2LjxAgAABAgQIECBAgJ4CTfNi95JfjW8FCBAgQJld6yWuu3VuNO0AAAAASUVORK5CYII='

export async function createDemoWorkbook(univerAPI) {
  const workbook = univerAPI.createWorkbook({ name: 'Univer 功能示例' })
  const sheet = workbook.getActiveSheet()
  sheet.setName('销售明细')

  // 标题（合并单元格）和表头
  sheet.getRange('A1:F1').merge()
  sheet.getRange('A1').setValue('2026 年 9 月销售明细')
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
  sheet.getRange('A2:F2').setValues([['日期', '区域', '产品', '单价', '数量', '金额']])
  sheet.getRange('A2:F2').setFontWeight('bold').setBackgroundColor('#DCE6F1')

  const rows = [
    ['2026-09-01', '华东', '苹果', 5.5, 120],
    ['2026-09-02', '华南', '香蕉', 3.2, 200],
    ['2026-09-03', '华北', '橙子', 4.8, 150],
    ['2026-09-04', '华东', '葡萄', 12, 60],
    ['2026-09-05', '西南', '西瓜', 2.5, 300],
    ['2026-09-06', '华南', '芒果', 15, 45],
    ['2026-09-07', '华北', '苹果', 5.8, 110],
    ['2026-09-08', '西南', '香蕉', 3, 260],
  ]
  const toSerial = (s) => (Date.UTC(...s.split('-').map((n, i) => (i === 1 ? Number(n) - 1 : Number(n)))) - Date.UTC(1899, 11, 30)) / 86400000
  sheet.getRange('A3:E10').setValues(rows.map(([d, ...rest]) => [toSerial(d), ...rest]))
  sheet.getRange('A3:A10').setNumberFormat('yyyy-mm-dd')
  sheet.getRange('D3:D10').setNumberFormat('0.00')
  for (let r = 3; r <= 10; r++) sheet.getRange(`F${r}`).setValue({ f: `=D${r}*E${r}` })
  sheet.getRange('F3:F11').setNumberFormat('#,##0.00')
  sheet.getRange('A11').setValue('合计')
  sheet.getRange('E11').setValue({ f: '=SUM(E3:E10)' })
  sheet.getRange('F11').setValue({ f: '=SUM(F3:F10)' })
  sheet.getRange('A11:F11').setFontWeight('bold')
  sheet.getRange('A2:F11').setBorder(univerAPI.Enum.BorderType.ALL, univerAPI.Enum.BorderStyleTypes.THIN, '#A6A6A6')
  sheet.setColumnWidth(0, 100)
  sheet.setFrozenRows(2)

  // 条件格式：单价 > 10 高亮、数量色阶、金额数据条
  sheet.addConditionalFormattingRule(sheet.getRange('D3:D10').createConditionalFormattingRule()
    .whenNumberGreaterThan(10).setBackground('#FFC7CE').setFontColor('#9C0006').build())
  sheet.addConditionalFormattingRule(sheet.getRange('E3:E10').createConditionalFormattingRule()
    .setColorScale([{ index: 0, color: '#F8696B', value: { type: 'min' } }, { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } }, { index: 2, color: '#63BE7B', value: { type: 'max' } }]).build())
  sheet.addConditionalFormattingRule(sheet.getRange('F3:F10').createConditionalFormattingRule()
    .setDataBar({ min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6', nativeColor: '#FF0000', isGradient: true, isShowValue: true }).build())

  // 数据验证：区域下拉、数量必须是 1~1000 的整数
  sheet.getRange('B3:B10').setDataValidation(univerAPI.newDataValidation().requireValueInList(['华东', '华南', '华北', '西南']).build())
  sheet.getRange('E3:E10').setDataValidation(univerAPI.newDataValidation().requireNumberBetween(1, 1000, true)
    .setOptions({ showErrorMessage: true, error: '数量必须是 1~1000 的整数', showInputMessage: true, promptTitle: '数量', prompt: '填写 1~1000 的整数' }).build())

  // 筛选
  sheet.getRange('A2:F10').createFilter()

  // 右侧：动态数组、LET、定义名称、超链接、富文本
  sheet.getRange('H2').setValue('区域（去重排序）')
  sheet.getRange('H3').setValue({ f: '=SORT(UNIQUE(B3:B10))' })
  sheet.getRange('J2').setValue('含税合计（LET）')
  sheet.getRange('J3').setValue({ f: '=LET(total,SUM(F3:F10),total*(1+税率))' })
  sheet.getRange('J3').setNumberFormat('#,##0.00')
  workbook.insertDefinedName('税率', '=0.13')
  workbook.insertDefinedName('单价列', '销售明细!$D$3:$D$10')
  sheet.getRange('J5').setValue('最高单价（定义名称）')
  sheet.getRange('J6').setValue({ f: '=MAX(单价列)' })
  sheet.getRange('J8').setValue('LAMBDA')
  sheet.getRange('J9').setValue({ f: '=LAMBDA(x,y,x*y)(D3,E3)' })
  sheet.setColumnWidth(7, 120)
  sheet.setColumnWidth(9, 130)

  try {
    await sheet.getRange('H9').setHyperLink('https://docs.univer.ai', 'Univer 文档')
    const rich = univerAPI.newRichText().insertText('红色加粗 普通文字')
    rich.setStyle(0, 4, univerAPI.newTextStyle().setColor({ rgb: '#FF0000' }).setBold(true).build())
    sheet.getRange('H11').setRichTextValueForCell(rich)
  } catch (err) { console.warn('[demo] 超链接/富文本', err) }

  // 批注和评论
  sheet.getRange('F11').createOrUpdateNote({ note: '合计金额 = 各行单价 × 数量之和', width: 200, height: 60, show: false })
  try {
    await sheet.getRange('C8').addCommentAsync(univerAPI.newTheadComment().setContent(univerAPI.newRichText().insertText('芒果单价偏高，确认一下进货价？')))
  } catch (err) { console.warn('[demo] 评论', err) }

  // 浮动图片
  try { await sheet.insertImage(CHART_PNG, 7, 12) } catch (err) { console.warn('[demo] 图片', err) }

  // 第二个工作表：超级表 + 区域保护
  const tableSheet = workbook.insertSheet('超级表')
  tableSheet.getRange('A1:D6').setValues([
    ['员工', '部门', '入职年份', '绩效'],
    ['张三', '销售', 2021, 'A'],
    ['李四', '研发', 2019, 'B'],
    ['王五', '销售', 2023, 'A'],
    ['赵六', '财务', 2020, 'C'],
    ['孙七', '研发', 2022, 'B'],
  ])
  try { await tableSheet.addTable('员工表', { startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 }) } catch (err) { console.warn('[demo] 超级表', err) }
  tableSheet.getRange('F1').setValue('F1:G3 是受保护区域')
  tableSheet.getRange('F2:G3').setValues([['受保护', '不能改'], ['受保护', '不能改']])
  try { await tableSheet.getRange('F2:G3').getRangePermission().protect({ name: '示例保护区域' }) } catch (err) { console.warn('[demo] 保护', err) }

  workbook.setActiveSheet(sheet)
  return workbook
}
