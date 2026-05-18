# bruce-svg-to-pptx

用 AI 生成 SVG，再自动转成 PowerPoint 原生可编辑形状——一套让 AI 直接产出可编辑演示文稿的工作流。

## 为什么用 SVG 作为中间格式

AI 不能直接操作 PowerPoint 内部格式（DrawingML），但非常擅长生成 SVG——SVG 本质上是结构化的矢量描述，和 AI 擅长输出的 XML/代码形式天然契合。

本工作流的思路：

```
AI 生成 SVG（图表、插图、布局）
        ↓
PowerPoint 导入 SVG，执行"转换为形状"
        ↓
原生 DrawingML 形状：可取消组合、可改色、可动画、完全可编辑
```

SVG → 形状的转换通过 PowerPoint 内置的 `ExecuteMso("SVGEdit")` 命令完成（即右键菜单里的"转换为形状"），产出结果与在 PowerPoint 里手动绘制的形状完全一致。

## 两种场景

### 场景一：从零生成演示文稿

AI 为每张幻灯片生成一个 SVG（内容 + 文字全部在 SVG 里布局），脚本批量转换为可编辑形状，产出一份完整的 .pptx。

```powershell
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath   .\ai-generated\ `
    -OutputPath .\deck.pptx -Force
```

### 场景二：在现有品牌模板上填充 AI 内容

公司模板的封面、目录、品牌设计保持不动；AI 只生成内容区的插图，脚本将其嵌入模板的安全内容区。

```
封面 / 目录页    →  只改文字（TEXT 模式）
内容页          →  复制模板页 + 嵌入 AI 生成的 SVG（EXPAND 模式）
整份演示文稿     →  用一个 JSON 描述所有页面，一次调用产出（MANIFEST 模式）
```

## 快速上手

### 新建演示文稿

```powershell
# 单个 SVG → 新建演示文稿
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\slide.svg -OutputPath .\out.pptx -Force

# 文件夹批量转换，每个 SVG 一张幻灯片
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\slides\ -OutputPath .\deck.pptx -Force
```

### 编辑品牌模板：只改封面/目录文字

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx -OutputPath .\out.pptx -Force `
    -SlideTexts @{
        1 = @{ Title = '2026 战略报告'; Subtitle = '董事会汇报'; Date = '2026-05' }
        2 = @{ Title = '目录'; Body = @('市场概述', '竞争格局', '战略规划', '执行路径') }
    }
```

### 编辑品牌模板：批量插入 AI 生成的内容页

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath     .\branded.pptx -OutputPath .\out.pptx -Force `
    -SvgPath          .\ai-generated\ `
    -ContentSlide     3 `
    -InsertAfterSlide 2 `
    -ClearContent `
    -SlideTexts @{
        1 = @{ Title = '第一章：市场概述' }
        2 = @{ Title = '第二章：竞争格局' }
    }
```

### 一次产出整份演示文稿（推荐）

用 JSON manifest 描述所有页面，避免复杂的命令行拼接：

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -Manifest     .\deck.json `
    -OutputPath   .\out.pptx -Force
```

Manifest 示例见 [`examples/deck.json`](examples/deck.json)，支持在同一文件里混合 `text`（改文字）、`expand`（插入内容页）、`insert`（替换指定页）三种操作。

## AI 生成 SVG 的注意事项

PowerPoint 的转换器只支持部分 SVG 特性。生成 SVG 时需遵守以下规则，否则转换后会变成一张图片而非可编辑形状：

**可用**

`<path>` `<rect>` `<circle>` `<ellipse>` `<line>` `<polygon>` `<polyline>`、`<text>`、纯色填充与描边、简单线性渐变（≤3 色标）、`<g>` 分组与基本变换（translate / scale / rotate）

**避免**

`<filter>`（模糊/投影）、`<clipPath>` `<mask>`、`<foreignObject>`、`<image>` 嵌入位图、径向渐变、`<style>` 块内的 CSS（改用内联 style 属性）、Web 字体（改用 Arial、Calibri、微软雅黑等系统字体）

**在品牌模板场景下**，AI 生成的 SVG 只负责内容安全区内的插图，**不要**在 SVG 里重绘模板的标题栏、Logo、页脚、背景——这些由模板幻灯片本身提供。

## 前置条件

- Windows 10 / 11
- Microsoft PowerPoint 2016 build 1712+（Microsoft 365 / Office 2019 / 2021 / 2024 均可）
- PowerShell 5.1（系统内置）或 PowerShell 7（`pwsh`，推荐，中文字符更稳定）

## 参数速查

| 参数 | 适用脚本 | 说明 |
|---|---|---|
| `-Force` | 两者 | **自动化调用时必须加**，跳过"是否覆盖"的交互确认 |
| `-ClearContent` | Edit | 插入前清除内容区内的非结构性形状 |
| `-ContentZone` | Edit | 手动指定内容区 `"Left,Top,Width,Height"`（单位：点，1英寸=72点） |
| `-NoBackup` | Edit | 跳过自动创建 `.bak.pptx` 备份 |
| `-Append` | Convert | 追加到已有演示文稿而非覆盖 |
| `-Manifest` | Edit | 指定 JSON manifest 文件，推荐用于复杂多页任务 |

## 验证结果

打开生成的 `.pptx`，右键单击插入的形状：

- 显示**编辑顶点**或可取消组合 → 转换成功，是真正的可编辑形状
- 仍显示**转换为形状** → 该 SVG 包含不支持的特性，查看脚本 stderr 输出定位原因

## 局限

- PowerPoint 窗口在转换期间必须可见（`ExecuteMso` 的硬性要求）
- 不支持无头 / CI 环境
- 每张幻灯片一个 SVG，不支持单页多 SVG 布局
- 仅支持 Windows
