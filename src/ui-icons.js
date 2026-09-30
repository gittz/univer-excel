// 自定义工具栏图标（Univer 自带图标库里没有文件 / 导入 / 导出 / 分列 / 水印图标）
// 用 Univer 组件管理器提供的 React createElement 画 SVG，不需要额外依赖 React。
import { IconManager } from '@univerjs/preset-sheets-core'

const PATHS = {
  // 文件：折角的文档
  CustomFileIcon: [
    ['path', { d: 'M4 1.75h5.25L12.5 5v9.25H4z' }],
    ['path', { d: 'M9.25 1.75V5h3.25' }],
    ['path', { d: 'M6 8.25h4.5M6 10.75h4.5' }],
  ],
  // 导入：箭头进托盘
  CustomImportIcon: [
    ['path', { d: 'M8 1.75v8' }],
    ['path', { d: 'M4.75 6.5 8 9.75l3.25-3.25' }],
    ['path', { d: 'M2.25 10.5v3.75h11.5V10.5' }],
  ],
  // 导出：箭头出托盘
  CustomExportIcon: [
    ['path', { d: 'M8 10V2.25' }],
    ['path', { d: 'M4.75 5.5 8 2.25l3.25 3.25' }],
    ['path', { d: 'M2.25 10.5v3.75h11.5V10.5' }],
  ],
  // 分列：一列拆成两列
  CustomSplitColumnsIcon: [
    ['rect', { x: 1.75, y: 2.25, width: 12.5, height: 11.5, rx: 1 }],
    ['path', { d: 'M8 2.25v11.5', strokeDasharray: '1.6 1.4' }],
    ['path', { d: 'M3.75 8h2.5M5.25 6.5 6.75 8l-1.5 1.5M12.25 8h-2.5M10.75 6.5 9.25 8l1.5 1.5' }],
  ],
  // 水印：倾斜的字母
  CustomWatermarkIcon: [
    ['rect', { x: 1.75, y: 1.75, width: 12.5, height: 12.5, rx: 1.5 }],
    ['path', { d: 'm4.5 11 2-6 1.5 4 1.5-4 2 6', transform: 'rotate(-12 8 8)' }],
  ],
}

export function registerCustomIcons(univer, univerAPI) {
  const { createElement } = univerAPI.getComponentManager().reactUtils
  const iconManager = univer.__getInjector().get(IconManager)
  const icons = {}
  for (const [name, shapes] of Object.entries(PATHS)) {
    const Icon = (props = {}) => createElement(
      'svg',
      {
        viewBox: '0 0 16 16',
        width: '1em',
        height: '1em',
        fill: 'none',
        stroke: 'currentColor',
        strokeWidth: 1.2,
        strokeLinecap: 'round',
        strokeLinejoin: 'round',
        className: props.className,
        style: props.style,
        'aria-hidden': true,
      },
      ...shapes.map(([tag, attrs], i) => createElement(tag, { key: i, ...attrs })),
    )
    Icon.displayName = name
    icons[name] = Icon
  }
  iconManager.register(icons)
}
