# Univer 开源版全功能示例

基于 Univer 开源版（Presets 方式）的电子表格，接入了开源版的全部表格功能，并实现了纯前端的 Excel 导入 / 导出。
**不使用任何 `@univerjs-pro` 包，也不需要服务端。**

## 运行

需要 Node.js 20.19+ 或 22.12+。

```bash
npm install
npm run dev
```

浏览器打开终端里显示的地址（默认 http://localhost:5173）。页面会打开一个示例工作簿，把下面的功能都用了一遍。

## 部署到 Cloudflare Pages

纯静态网站，构建产物在 `dist/`（`index.html` + `assets/` + `_headers`）。

**方式一：直接上传**
1. `npm run build`（或直接用打包好的 `univer-excel-dist.zip`）
2. Cloudflare 控制台 → Workers & Pages → 创建 → Pages → 上传资产
3. 把 `univer-excel-dist.zip` 拖进去（或选择 `dist` 文件夹），部署

**方式二：命令行**
```bash
npm run build
npx wrangler pages deploy dist --project-name univer-excel
```

**方式三：连接 GitHub 自动部署（推荐）**

推送到 GitHub 的 `main` 分支后，Cloudflare 会自动安装依赖、构建并发布；其他分支会生成预览地址。

Cloudflare 控制台 → Workers 和 Pages → 创建 → Pages → 导入现有 Git 存储库 → 选择本仓库，构建设置：

| 项目 | 值 |
| --- | --- |
| 生产分支 | `main` |
| 框架预设 | 无（或 Vite） |
| 构建命令 | `pnpm run build` |
| 构建输出目录 | `dist` |
| 环境变量 | `NODE_VERSION` = `22`，`PNPM_VERSION` = `9` |

如果在 Cloudflare 里建的是 **Workers** 项目（导入 Git 仓库时的默认选项），用仓库里的 `wrangler.jsonc`，构建设置：

| 项目 | 值 |
| --- | --- |
| 构建命令 | `pnpm run build` |
| 部署命令 | `npx wrangler deploy` |
| 构建变量 | `PNPM_VERSION` = `9` |

`wrangler.jsonc` 声明把 `dist` 作为静态资源发布；没有这个文件时 `wrangler deploy` 会尝试自动改 `vite.config.js` 并报错 “Cannot modify Vite config: could not find a valid plugins array”。

仓库里只保留 `pnpm-lock.yaml` 一个锁文件（`package-lock.json` 已加入 `.gitignore`），`.node-version` 固定 Node 22。

说明：
- `vite.config.js` 里 `base: './'`，部署在域名根目录或子路径下都能用。
- `public/_headers` 让带哈希的 `assets/*` 长期缓存，`index.html` 每次校验，更新后用户刷新就能拿到新版本。
- 最大的单个文件约 9 MB，在 Cloudflare Pages 单文件 25 MiB 的限制内；Cloudflare 会自动压缩传输。

## 已接入的开源功能

| 领域 | 功能 | 入口 |
| --- | --- | --- |
| 单元格与样式 | 值和类型、富文本、字体、填充、边框、对齐、换行、旋转、合并 | 开始 |
| 数字格式 | Excel 格式串 | 开始 |
| 公式 | 约 510 个函数、动态数组、LAMBDA/LET、跨表引用、定义名称 | 公式 |
| 表格操作 | 行高列宽、隐藏、冻结、网格线、自动填充、格式刷、剪贴板、分列 | 开始 / 视图 / 数据 → 分列 |
| 数据 | 排序、筛选、数据验证、条件格式、超级表 | 数据 |
| 协作相关 | 批注、评论、撤销重做、工作表和区域保护 | 插入 / 右键 / 开始 → 保护 |
| 其他 | 查找替换、超链接、浮动图片和单元格内图片、水印 | 数据 / 插入 / 视图 → 水印 |

形状（shapes）只在 Univer Pro 里有，开源版没有，所以没有接入。

## Excel 导入 / 导出

工具栏最左侧的「导入 Excel」「导出 Excel」。基于开源的 ExcelJS，ExcelJS 处理不了的部分直接读写 xlsx 里的 XML（JSZip）。

| 功能 | 导出 | 导入 | 说明 |
| --- | --- | --- | --- |
| 值、公式、样式、数字格式、合并、行高列宽、隐藏、冻结、网格线、标签颜色 | ✓ | ✓ | |
| 富文本 | ✓ | ✓ | |
| 新函数（XLOOKUP、FILTER、LET、LAMBDA 等） | ✓ | ✓ | 自动加 / 去 Excel 要求的 `_xlfn.` `_xlpm.` 前缀 |
| 动态数组 | ≈ | ✓ | 导出为数组公式 `{=...}`，结果一致 |
| 定义名称 | ✓ | ✓ | 支持公式型名称、工作表级作用域 |
| 超链接 | ✓ | ✓ | 含工作表内部跳转 |
| 条件格式 | ✓ | ✓ | 高亮（数值 / 文本 / 日期 / 排名 / 平均值 / 重复 / 唯一 / 公式）、色阶、数据条、图标集 |
| 数据验证 | ✓ | ✓ | 下拉（列表 / 区域）、整数、小数、文本长度、日期、时间、自定义公式、提示和出错信息；复选框导出为 TRUE/FALSE 下拉 |
| 筛选 | ✓ | ✓ | 含按值筛选、自定义条件筛选和已筛掉的行；按颜色筛选暂不支持 |
| 超级表 | ✓ | ✓ | |
| 批注 | ✓ | ✓ | |
| 评论 | ≈ | ✓ | 导入支持 Excel 365 新式评论（含作者和回复）；导出为批注（含作者、时间、回复） |
| 浮动图片 | ✓ | ✓ | png / jpeg / gif |
| 单元格内图片 | ≈ | — | 导出为贴在单元格上的浮动图片 |
| 工作表保护 / 区域保护 | ✓ / ≈ | ✓ | 区域保护导出为“保护工作表 + 其余单元格解除锁定”；Excel 的“保护工作表 + 部分单元格未锁定”导入为“未锁定区域以外全部受保护”（含空单元格） |
| 水印 | — | — | 显示层效果，Excel 没有对应概念 |

**保护按 Excel 的规则生效：** Univer 开源版默认是“创建者仍可编辑”，本项目改成和 Excel 一样——受保护的单元格谁都不能改（包括创建者），在保护面板里删除保护后才能编辑。导入的文件、示例工作簿和界面上新建的保护都一样（`src/protection.js`）。

不支持旧版 `.xls` 文件。导入导出时遇到近似处理或不支持的内容，会在浏览器控制台里输出提示。

## 文件

- `vite.config.js`：构建配置（相对路径）
- `.node-version`：Cloudflare 构建使用的 Node 版本
- `wrangler.jsonc`：Cloudflare Workers 部署配置（把 `dist` 作为静态网站发布）
- `public/_headers`：Cloudflare Pages 缓存规则
- `src/main.js`：创建 Univer，接入各功能预设，添加导入 / 导出 / 分列 / 水印按钮
- `src/demo-data.js`：示例工作簿
- `src/protection.js`：让保护按 Excel 规则生效
- `src/export-xlsx.js`、`src/import-xlsx.js`：导出 / 导入主流程
- `src/xlsx/utils.js`：坐标、颜色、样式、尺寸换算
- `src/xlsx/formula.js`：Excel 新函数前缀处理
- `src/xlsx/conditional-format.js`：条件格式互转
- `src/xlsx/data-validation.js`：数据验证互转
- `src/xlsx/xml-parts.js`：直接读写 xlsx XML（定义名称、筛选条件、数据验证、新式评论、兼容性修正）

打开页面后可在浏览器控制台里用 `univerAPI` 调用 Facade API。
