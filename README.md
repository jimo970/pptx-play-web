# Deckline PPTX 播放器

Deckline 是一个在浏览器中打开和播放 PowerPoint 演示文稿的应用。选择或拖入未加密的 `.pptx` 文件即可开始；文件解析在当前浏览器中完成，项目没有演示文稿上传接口。

## 功能

- 浏览幻灯片缩略图，切换页面，并播放支持的动画和切换效果。
- 使用幻灯片浏览视图整理页面，或在阅读视图中查看内容与备注。
- 在演讲者视图中查看当前页、下一页、演讲者备注和计时器；可打开同步的第二屏幕窗口。
- 放映时使用指针、画笔、荧光笔和激光笔，并可切换黑屏或白屏。
- 支持键盘和触屏操作：触屏设备可滑动翻页、双指缩放；放大后可拖动查看。
- 界面支持中文和英文。

## 快速开始

需要可运行 Node.js 项目的环境。

```sh
npm install
npm run dev
```

开发服务器默认地址为 <http://127.0.0.1:3000>。启动后，点击“打开 PPTX”或把文件拖到页面中。

## 常用快捷键

- `←` / `→`：上一步或下一步动画；没有待播放动画时切换幻灯片。
- `F5`：从头开始放映；`Shift+F5`：从当前页开始放映。
- `F`：切换全屏；`G`：打开幻灯片浏览视图；`S`：打开演讲者视图。
- 演讲者视图的第二屏幕窗口支持方向键翻页，`B` / `W` 可切换黑屏或白屏，`Esc` 可关闭。

## 文件兼容性

应用读取 `.pptx`（Office Open XML）文件，不支持旧版 `.ppt` 文件或带密码的加密演示文稿。浏览器输入限制包括：文件不超过 150 MB、解压后不超过 250 MB、压缩包条目不超过 20,000 个、幻灯片不超过 500 张。

文本、图形、图片、表格、图表、音视频、备注、超链接、动画和切换效果的呈现取决于文件使用的 OOXML 特性。对于不支持或只能近似呈现的内容，播放器可能显示兼容性提示，或跳过相关效果。浏览器渲染与 PowerPoint / WPS 的视觉和播放结果不保证完全一致。

详细的兼容性实现清单、验证范围和已知限制见[英文版 README](./README.en.md)。

## 开发命令

```sh
npm run dev       # 启动开发服务器
npm run typecheck # TypeScript 类型检查
npm run build     # 构建生产版本
npm run preview   # 预览生产构建
```

仓库还提供以下专项校验脚本：

| 命令 | 用途 |
| --- | --- |
| `npm run verify:animation` | 在本地 Edge 或 Chrome 中校验动画播放；需要开发服务器运行。 |
| `npm run verify:office-animation-playback` | 校验 Office / WPS 导出的动画文件；按脚本要求提供 PPTX fixture 和 evidence 目录。 |
| `npm run verify:chart-number-format` | 校验图表数字格式处理。 |
| `npm run verify:narration` | 运行旁白相关的动画校验。 |
| `npm run verify:pptx-input-errors` | 校验 PPTX 输入错误处理。 |

## 项目目录

| 路径 | 说明 |
| --- | --- |
| `app/pages/index.vue` | 播放器主界面和放映逻辑。 |
| `app/pages/projector.vue` | 第二屏幕投影页面。 |
| `app/components/SlideStage.vue` | 幻灯片内容渲染。 |
| `app/utils/pptx.ts` | PPTX 解析与模型构建。 |
| `app/workers/pptx-parser.worker.ts` | 在 Web Worker 中执行解析。 |
| `i18n/locales/` | 中文与英文界面文案。 |
| `scripts/` | 自动化校验与 fixture 工具。 |

## 技术栈

Nuxt 4、Vue 3、TypeScript、Nuxt i18n；使用 JSZip 读取 PPTX 压缩包，并通过 Web Worker 隔离解析工作。

